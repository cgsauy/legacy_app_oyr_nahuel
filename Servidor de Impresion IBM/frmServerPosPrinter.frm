VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmServerPosPrinter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Pos Printer"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   6705
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServerPosPrinter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmCliente 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2520
      Top             =   2280
   End
   Begin VB.Timer tmStatus 
      Left            =   400
      Top             =   2280
   End
   Begin VB.Timer tmLectura 
      Enabled         =   0   'False
      Left            =   1800
      Top             =   2280
   End
   Begin MSCommLib.MSComm mscPOS 
      Left            =   960
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton butActivarPos 
      Caption         =   "&Activar servidor"
      Height          =   435
      Left            =   4560
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   6675
      TabIndex        =   2
      Top             =   0
      Width           =   6705
      Begin VB.Image imgPrint 
         Height          =   480
         Left            =   120
         Picture         =   "frmServerPosPrinter.frx":1708A
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Impresión de documentos POS Printer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.ComboBox cboPos 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Text            =   "cboPos"
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Label lblActividad 
      BackStyle       =   0  'Transparent
      Caption         =   "lblActividad"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Label lblRollo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rollo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000060&
      Height          =   855
      Left            =   4440
      TabIndex        =   5
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione el nombre del pos al cual se dirigen los documentos"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   3615
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu MnuArcEstado 
         Caption         =   "Estado impresora"
      End
      Begin VB.Menu MnuReimprimir 
         Caption         =   "Reimprimir"
         Begin VB.Menu MnuRePrintNroDoc 
            Caption         =   "Un ticket"
         End
         Begin VB.Menu MnuRePrintAPartirDe 
            Caption         =   "Muchos tickets (rango desde y hasta)"
         End
         Begin VB.Menu MnuLineTickets 
            Caption         =   "-"
         End
         Begin VB.Menu MnuTicketCI 
            Caption         =   "Un ticket por C.I."
         End
         Begin VB.Menu MnuTicketsRangoCI 
            Caption         =   "Rango de C.I."
         End
      End
      Begin VB.Menu MnuArchLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuInformeZ 
         Caption         =   "Informe Z (cierre del día)"
      End
      Begin VB.Menu MnuArchExit 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu MnuConfiguracion 
      Caption         =   "Configuración"
      Begin VB.Menu MnuCnfTickeadoraRP 
         Caption         =   "Soy tickeadora de pagos por Giros"
      End
      Begin VB.Menu MnuCnfLine0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuArchCnfgPuerto 
         Caption         =   "Configurar puerto"
      End
      Begin VB.Menu MnuArcLimpiarMemoria 
         Caption         =   "Limpiar memoria"
      End
      Begin VB.Menu MnuArcGrabarlogo 
         Caption         =   "Almacenar Logo"
      End
      Begin VB.Menu MnuArcGrabarLinea 
         Caption         =   "Almacenar línea"
      End
      Begin VB.Menu MnuArcLine1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuArcLogoAlmacenado 
         Caption         =   "Imprimir logo almacenado"
      End
      Begin VB.Menu MnuArcImprimirLogo 
         Caption         =   "Imprimir Logo"
      End
      Begin VB.Menu MnuArcImprimirLinea 
         Caption         =   "Imprimir línea"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmServerPosPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim serieRollo As String
Dim miPrinter As Long

Dim colCR As Collection
Dim colJournal As Collection

Dim EstadoImpresora As Byte

Dim oPos As New clsPosIBM4610

Dim sBufferMsgPrint As String   'buffer de respuesta de la pos.
Dim iStatus As Byte

Private colDocumentosAImprimir As Collection

Private Sub BuscoCodigoEnCombo(cCombo As Control, lngCodigo As Long)
Dim i As Integer
    
    If cCombo.ListCount > 0 Then
        For i = 0 To cCombo.ListCount - 1
            If cCombo.ItemData(i) = lngCodigo Then
                cCombo.ListIndex = i
                Exit Sub
            End If
        Next i
        cCombo.ListIndex = -1
    Else
        cCombo.ListIndex = -1
    End If

End Sub

Function SeleccionarMiImpresoraAsignada() As Boolean

    miPrinter = 0
    'IDTerminal:IDTickeadora|...|idTn:IDTn'
    If prmTickeadorasAsignadas <> "" Then
        
        Dim vTrama() As String
        Dim vValores() As String
        vTrama = Split(prmTickeadorasAsignadas, "|")
        Dim iQ As Integer
        For iQ = 0 To UBound(vTrama)
            If vTrama(iQ) <> "" Then
                vValores = Split(vTrama(iQ), ":")
                If Val(vValores(0)) = paCodigoDeTerminal Then
                    miPrinter = Val(vValores(1))
                    BuscoCodigoEnCombo cboPos, Val(vValores(1))
                    Exit Function
                End If
            End If
        Next
        
    End If
    
    MsgBox "Su terminal no tiene asignada ninguna impresora. " & vbCrLf & vbCrLf & _
        "Debe seleccionar la impresora y en el momento que de ACTIVAR quedará asignada a su Terminal.", vbExclamation, "ATENCIÓN"
   
    
End Function

Function BuscarTerminalPorImpresora(ByVal idP As Integer) As String
Dim idTer As Integer

    Dim vTrama() As String
    Dim vValores() As String
    vTrama = Split(prmTickeadorasAsignadas, "|")
    Dim iQ As Integer
    For iQ = 0 To UBound(vTrama)
        If vTrama(iQ) <> "" Then
            vValores = Split(vTrama(iQ), ":")
            If Val(vValores(1)) = idP Then
                idTer = Val(vValores(0))
                Exit For
            End If
        End If
    Next
    
    If idTer > 0 Then
        Cons = "SELECT TerNombre FROM Terminal WHERE TerCodigo = " & idTer
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then BuscarTerminalPorImpresora = Trim(RsAux(0))
        RsAux.Close
    End If
    
End Function

Sub GuardoSeteoImpresora()

Dim nomTer As String
            
    CargarImpresorasAsignadas
    
    If cboPos.ListIndex = -1 Then
        MsgBox "Atención su terminal a partir de este momento no posee tickeadora asignada.", vbInformation, "ATENCIÓN"
    Else
        nomTer = BuscarTerminalPorImpresora(cboPos.ItemData(cboPos.ListIndex))
        If nomTer <> "" Then
            If MsgBox("La tickeadora seleccionada pertenece a la terminal " & nomTer & vbCrLf & vbCrLf & _
                    "¿Confirma asignar la tickeadora seleccionada a su terminal de todas formas?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
                BuscoCodigoEnCombo cboPos, miPrinter
                Exit Sub
            End If
        ElseIf MsgBox("¿Confirma asignar la tickeadora seleccionada a su terminal?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
            BuscoCodigoEnCombo cboPos, miPrinter
            Exit Sub
        End If
    End If
    
    
    Dim colTerTic As New Collection
    Dim oTerTic As New clsTerminalTickeadora

    Dim salida As String
    Dim bTomeAccion As Boolean
    
    Dim mSeleccionada As Long
    If cboPos.ListIndex > -1 Then mSeleccionada = cboPos.ItemData(cboPos.ListIndex)
    
    If (prmTickeadorasAsignadas = "") Then
        
        salida = paCodigoDeTerminal & ":" & mSeleccionada
        miPrinter = mSeleccionada
        bTomeAccion = True
        
    Else
    
        Dim vTrama() As String
        Dim vValores() As String
        vTrama = Split(prmTickeadorasAsignadas, "|")
        
        Dim iQ As Integer
        For iQ = 0 To UBound(vTrama)
            
            If vTrama(iQ) <> "" Then
                vValores = Split(vTrama(iQ), ":")
                
                
                oTerTic.Terminal = CLng(vValores(0))
                oTerTic.Tickeadora = CLng(vValores(1))
                
                If oTerTic.Terminal = paCodigoDeTerminal Then
                    
                    If bTomeAccion Then
                        'ya hice el cambio por la impresora así que borro este item.
                        oTerTic.Terminal = 0
                    Else
                        If cboPos.ListIndex = -1 Then
                            'Se está quitando.
                            oTerTic.Tickeadora = 0
                            miPrinter = 0
                            
                        ElseIf oTerTic.Tickeadora <> mSeleccionada Then
                            
                            oTerTic.Tickeadora = mSeleccionada
                            miPrinter = oTerTic.Tickeadora
                            
                        End If
                        bTomeAccion = True
                        
                    End If
                    
                Else
                
                    If oTerTic.Tickeadora = mSeleccionada Then
                        If bTomeAccion Then
                            oTerTic.Tickeadora = 0
                        Else
                            bTomeAccion = True
                            oTerTic.Terminal = paCodigoDeTerminal
                        End If
                    
                        
                    End If
                    
                End If
                
                If oTerTic.Tickeadora > 0 And oTerTic.Terminal > 0 Then
                    
                    salida = salida & IIf(salida <> "", "|", "") & oTerTic.Terminal & ":" & oTerTic.Tickeadora
                    
                End If
            End If
            
        Next
        
    End If
    
    If Not bTomeAccion And cboPos.ListIndex > -1 Then
        salida = salida & IIf(salida <> "", "|", "") & paCodigoDeTerminal & ":" & mSeleccionada
        miPrinter = mSeleccionada
        bTomeAccion = True
    End If
    
    If bTomeAccion Then
        Cons = "UPDATE Parametro SET ParTexto = '" & salida & "' WHERE ParNombre = 'TickeadoraAsignadaTerminal'"
        cBase.Execute (Cons)
    End If

End Sub

Function ValidoCierreFechaAnterior() As Boolean
On Error GoTo errVCFA

    Dim sPos As String
    Cons = "SELECT TicNombre FROM Tickeadora WHERE TicID = " & cboPos.ItemData(cboPos.ListIndex)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    sPos = ""
    If Not RsAux.EOF Then
        sPos = Trim(RsAux(0))
    End If
    
    'If Cons = "" Then Exit Function
    
    Dim fechaRecibo As Date
    fechaRecibo = DateSerial(1990, 1, 1)
    
    Cons = "SELECT Max(DocFecha) " _
        & "FROM Documento INNER JOIN TicketsAImprimir ON TAIDocumento = DocCodigo " _
        & "WHERE TAITickeadora = " & cboPos.ItemData(cboPos.ListIndex)
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux(0)) Then fechaRecibo = RsAux(0)
    RsAux.Close
    
    Dim sUltCierre As String
    
    sUltCierre = DateSerial(1990, 1, 1)
    
    Cons = "SELECT ParTexto From Parametro WHERE ParNombre = '" & sPos & "_UltCierre" & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux(0)) Then sUltCierre = Trim(RsAux(0))
    End If
    RsAux.Close
    
    If CDate(Format(fechaRecibo, "dd/mm/yyyy")) < Date And fechaRecibo > DateSerial(1990, 1, 1) And CDate(sUltCierre) < fechaRecibo Then
    
        If MsgBox("ATENCIÓN!!!!!" & vbCrLf & vbCrLf & vbCrLf & "No se realizó el INFORME Z del último día, ¿desea imprimirlo ahora?", vbQuestion + vbYesNo, "CIERRE DE TICKETS") = vbYes Then
        
            ApagarServidor
            ImprimoPrimerDocumento (True)
        
        End If
        
    End If
    
    
Exit Function
errVCFA:
    clsGeneral.OcurrioError "Error al validar "
End Function

Sub ApagarServidor()
    butActivarPos.Tag = 0
    tmLectura.Enabled = False
    tmStatus.Enabled = False
    butActivarPos.Enabled = True    'Si lo invoco para ir hacía atrás puede estar deshabilitado.
    sBufferMsgPrint = ""
    butActivarPos.Caption = "Activar POS"
    IndicarActividad "Servidor inactivo"
    cboPos.Enabled = True
    frmInicio.CambiarIcono "Servidor apagado", Error
End Sub

Public Sub EnviarAImpresora(ByVal txtaimpresora As String)
    mscPOS.Output = txtaimpresora
End Sub

Function ValidarNroRollo() As Boolean
Dim sRollo As String
Dim Seguir As Boolean
    
    If Val(lblRollo.Tag) = 0 Then
        ValidoCierreFechaAnterior
    End If
    
    Seguir = True
    Dim sSerie As String
    Do While Seguir
        sRollo = Trim(InputBox("Ingrese serie y número de rollo.", "Ingreso nro. rollo", serieRollo))
        If sRollo <> "" And sRollo <> serieRollo Then
        
            'Saco la serie.
            sSerie = Mid(sRollo, 1, 1)
            If Trim(serieRollo) <> "" And serieRollo <> sSerie Then
                If MsgBox("Cambió la serie del número en el rollo?" & vbCrLf & vbCrLf & "Presione SI para continuar o NO para cancelar", vbQuestion + vbYesNo + vbDefaultButton2, "Posible error") = vbNo Then Exit Function
            End If
            sRollo = Mid(sRollo, 2)
                        
            If Val(lblRollo.Tag) > 0 Then
                If Abs(Val(sRollo) - Val(lblRollo.Tag)) > 1 Then
                    MsgBox "El rollo siguiente sería el " & Val(lblRollo.Tag) + 1 & ", controle que el número ingresado es el correcto.", vbExclamation, "VALIDACIÓN"
                End If
            End If
            
            If MsgBox("¿Confirma que el número del rollo es: " & sRollo & " ?", vbQuestion + vbYesNo + vbDefaultButton2, "Nuevo rollo") <> vbYes Then
                ValidarNroRollo = False
                Exit Function
            End If
            
            lblRollo.Caption = "Rollo"
            If Val(sRollo) > 0 Then
                lblRollo.Caption = lblRollo.Caption & vbCrLf & Val(sRollo)
                lblRollo.Tag = Val(sRollo)
                serieRollo = sSerie
                Seguir = False
            Else
                lblRollo.Tag = 0
            End If
        Else
            Seguir = False
        End If
    Loop
    ValidarNroRollo = (Val(lblRollo.Tag) > 0)
    
End Function

Private Sub butActivarPos_Click()
    
    sBufferMsgPrint = ""
    EstadoImpresora = 0
    If Val(butActivarPos.Tag) = 1 Then
        ApagarServidor
    Else
        
        'Si cambió guardo la configuración.
        Dim mSel As Long
        If cboPos.ListIndex > -1 Then mSel = cboPos.ItemData(cboPos.ListIndex)
        
        If miPrinter <> mSel Then
            GuardoSeteoImpresora
        Else
            If cboPos.ListIndex = -1 Then MsgBox "Debe seleccionar una impresora para activar.", vbInformation, "ATENCIÓN": Exit Sub
        End If
        
        If miPrinter = 0 Or cboPos.ListIndex = -1 Then Exit Sub
        If miPrinter <> cboPos.ItemData(cboPos.ListIndex) Then Exit Sub
    
        If ValidarNroRollo Then
            cboPos.Enabled = False
            butActivarPos.Tag = 1
            butActivarPos.Caption = "Desactivar POS"
            tmLectura.Enabled = True
            tmLectura.Interval = 9900
        End If
    End If
    Me.Refresh
    
End Sub

Private Sub Form_Load()
On Error GoTo errL
    
    butActivarPos.Tag = "0"
    
    lblRollo.Tag = 0

    Dim sPort As String
    sPort = GetSetting(App.Title, "Configuración", "Puerto", sPort)
    If IsNumeric(sPort) Then sPort = Val(sPort) Else sPort = "1"
    
    cboPos.Clear
    CargoCombo "SELECT TicID, TicNombre From Tickeadora WHERE TicSucursal = " & paCodigoDeSucursal, cboPos, ""
    SeleccionarMiImpresoraAsignada
    
    If cboPos.ListIndex > -1 Then
        If cboPos.ItemData(cboPos.ListIndex) = prmTickeadoraCuotasGiros Then
            MnuCnfTickeadoraRP.Checked = True
        End If
    End If

    butActivarPos.Enabled = (cboPos.ListCount > 0)
    
    IndicarActividad "Abriendo puerto ..."
    mscPOS.InputLen = 1
    mscPOS.RThreshold = 1
    mscPOS.CommPort = Val(sPort)
    
    mscPOS.PortOpen = True
    
    IndicarActividad "Listo"
    Screen.MousePointer = 0
    Exit Sub
    
errL:
    clsGeneral.OcurrioError "Error al iniciar el formulario.", Err.Description & vbCrLf & "Puerto: " & Val(sPort), "Error"
    Screen.MousePointer = 0
End Sub

Sub IndicarActividad(ByVal msg As String)
    lblActividad.Caption = msg
    Me.Refresh
End Sub

Function BuscoDocumentosAImprimir() As Boolean
On Error GoTo errBDI
    
    IndicarActividad "Buscando documentos ..."
    
    Set colDocumentosAImprimir = Nothing
    Set colDocumentosAImprimir = New Collection
    Dim rsR As rdoResultset
    
    '[dbo].[prg_POS_TicketsAImprimir] @paso tinyint, @tickeadora smallint, @documento int = null, @nrorollo int = null
    Set rsR = cBase.OpenResultset("EXEC prg_POS_TicketsAImprimir 1, " & cboPos.ItemData(cboPos.ListIndex), rdOpenDynamic, rdConcurValues) '
    Do While Not rsR.EOF
        colDocumentosAImprimir.Add (rsR(0))
        EmitirCFE rsR(0)
        rsR.MoveNext
    Loop
    rsR.Close
    'colDocumentosAImprimir.Add (-369184)
    IndicarActividad "Proceso documentos ..."
    BuscoDocumentosAImprimir = (colDocumentosAImprimir.Count > 0)
    Exit Function
    
errBDI:
    BuscoDocumentosAImprimir = False
    clsGeneral.OcurrioError "Error al buscar los documentos pendientes de impresión.", Err.Description, "Impresión POS"
    Screen.MousePointer = 0
End Function

Function ObtenerRenglonesDelDocumentoAImprimir(ByVal idDoc As Long, ByVal accion As Byte) As Collection
    'voy al SP que me retorne los datos a imprimir.
    
    On Error GoTo errRD
    Screen.MousePointer = 11
    Dim retorno As New Collection
    Dim rsR As rdoResultset
    Dim oLinea As clsLineaImpresion
    'prg_POS_TicketsAImprimir @paso tinyint, @tickeadora smallint, @documento int = null
    Set rsR = cBase.OpenResultset("EXEC prg_POS_TicketsAImprimir " & accion & ", " & _
                    cboPos.ItemData(cboPos.ListIndex) & ", " & idDoc & ", " & Val(lblRollo.Tag) & ", '" & serieRollo & "'", rdOpenDynamic, rdConcurValues)
   ' Set rsR = cBase.OpenResultset("EXEC prg_POS_TicketsAImprimir 2, 1, " & idDoc, rdOpenDynamic, rdConcurValues)
    
    Do While Not rsR.EOF
        
        If (rsR(0) = -1) Then
            Exit Do
        End If
        
        Set oLinea = New clsLineaImpresion
        
        oLinea.Dato = rsR("Dato")
        If Not IsNull(rsR("Columnas")) Then oLinea.Columnas = rsR("Columnas")
        If Not IsNull(rsR("Fuente")) Then oLinea.Fuente = rsR("Fuente")
        oLinea.Recipiente = rsR("Recipiente")
        retorno.Add oLinea
        
        rsR.MoveNext
        
    Loop
    rsR.Close
    
    Set ObtenerRenglonesDelDocumentoAImprimir = retorno
    
    Screen.MousePointer = 0
    Exit Function
errRD:
    clsGeneral.OcurrioError "Error al cargar los renglones para el documento: " & idDoc, Err.Description, Me.Caption
End Function

Sub ImprimoPrimerDocumento(Optional ByVal EsCierre As Boolean = False)
On Error GoTo errID
Dim iCont As Integer
Dim oSalida As clsLineaSalida
Dim sPasoError As String

    sPasoError = "Tomo id documento"
    Dim doc As Variant
    If Not EsCierre Then
        If (colDocumentosAImprimir Is Nothing) Then
            ImprimoDocumentos
            Exit Sub
        ElseIf colDocumentosAImprimir.Count = 0 Then
            ImprimoDocumentos
            Exit Sub
        End If
    
        doc = CLng(colDocumentosAImprimir(1))
        IndicarActividad "Leyendo documento " & doc
    Else
        doc = 0
    End If
    
    'Leo el documento si este aún no fué impreso entonces me retorna datos.
    sPasoError = "leo el documento en BD"
    Dim colRenglonesImpresora As Collection
    Set colRenglonesImpresora = ObtenerRenglonesDelDocumentoAImprimir(doc, IIf(EsCierre, 4, 2))
    
    sPasoError = "Valido lectura"
    IndicarActividad "Imprimiendo documento " & doc
    If Not (colRenglonesImpresora Is Nothing) Then
    
        If colRenglonesImpresora.Count > 0 Then
        
            '1ero formo las lineas y luego envío a la impresora.
            
            Dim arrColumnas() As String
            Dim arrColFuente() As String
            Dim sFormatoUnico As String
            Dim arrDatos() As String
            Dim sLinea As String
            
            Set colCR = New Collection
            Set colJournal = New Collection
            
            Dim oLinea As clsLineaImpresion
            For Each oLinea In colRenglonesImpresora
                Erase arrColumnas
                Erase arrColFuente
                Erase arrDatos
                                    
                If oLinea.Columnas <> "" Then
                    arrColumnas = Split(oLinea.Columnas, "|")
                Else
                    ReDim arrColumnas(0)
                End If
                arrColFuente = Split(oLinea.Fuente, "|")
                
                arrDatos = Split(oLinea.Dato, "|")
                
                'Armo la linea
                sLinea = ""
                
                Dim sTabulador As String
                sTabulador = ""
                
                Dim oFuente As clsFuente
                Set oFuente = New clsFuente
                
                'Tomo la fuente.
                If UBound(arrColFuente) > 0 Then
                    
                    
                    oFuente.TipoFuente = Trim(arrColFuente(0))
                    If UBound(arrColFuente) >= 1 Then oFuente.High = (Val(arrColFuente(1)) = 1)
                    If UBound(arrColFuente) >= 2 Then oFuente.Wide = (Val(arrColFuente(2)) = 1)
                    If UBound(arrColFuente) >= 3 Then oFuente.Size = Val(arrColFuente(3))
                    If UBound(arrColFuente) >= 4 Then oFuente.Invert = (Val(arrColFuente(4)) = 1)
                
                    'If UBound(arrColFuente) >= 6 Then sTabulador = arrColFuente(6)
                
                End If
                
                If (UBound(arrColumnas) > 0) Then
                    'Ahora determino si es por tab o por cantidad de caracteres.
                    Dim bTab As Boolean: bTab = False
                    Dim arrColsCnfg() As String
                    For iCont = 0 To UBound(arrColumnas)
                        arrColsCnfg = Split(arrColumnas(iCont), ";")
                        If Val(arrColsCnfg(0)) > 0 Then bTab = True: Exit For
                    Next
                    
                    If bTab Then
                    
                        'Línea con TAB
                        For iCont = 0 To UBound(arrColumnas)
                            sLinea = sLinea & oPos.GetTxtHexadecimal(Val(arrColumnas(iCont)))
                        Next
                        sLinea = oPos.Key_SetTAB & sLinea & oPos.Key_SetTABEND
                        
                        'sLinea = sLinea & Join(arrDatos, oPos.key_TAB)
                         
                        'Armo la línea final
                        
                        'IMPORTANTE PRIMERO VA EL TAB Y LUEGO EL TEXTO.
                        For iCont = 0 To UBound(arrColumnas)
                            'Ya que el primero es tab ya lo seteo.
                            sLinea = sLinea & _
                                oPos.key_TAB & arrDatos(iCont)
                        Next
                    
                    Else
                    
                        Dim qCaract As Byte
                        Dim align As Byte
                        For iCont = 0 To UBound(arrColumnas)
                            Erase arrColsCnfg
                            arrColsCnfg = Split(arrColumnas(iCont), ";")
                            qCaract = Val(arrColsCnfg(1))
                            align = Val(arrColsCnfg(2))
                            
                            If Len(Trim(arrDatos(iCont))) < qCaract And align = 1 Then
                            
                                sLinea = sLinea & Space(qCaract - Len(Trim(arrDatos(iCont))))
                            
                            End If
                            sLinea = sLinea & arrDatos(iCont)
                            
                        Next
                    End If
                    
                Else
                
                    If UBound(arrDatos) = 0 Then
                        sLinea = arrDatos(0)
                    Else
                        sLinea = ""
                    End If
            
                End If
                
                Set oSalida = New clsLineaSalida
                oSalida.TextoSalida = sLinea
                Set oSalida.Fuente = oFuente
                oSalida.Tabulacion = sTabulador
                
                If oLinea.Recipiente = 0 Or oLinea.Recipiente = 1 Then colCR.Add oSalida
                If oLinea.Recipiente = 0 Or oLinea.Recipiente = 2 Then colJournal.Add oSalida
                
            Next
            
            butActivarPos.Enabled = False
            
            If EsCierre Then
                EnvioDatosAPos colJournal, False, Journal
                If colCR.Count > 0 Then
                    sPasoError = "Imprimo en cliente doc: " & doc
                    EnvioDatosAPos colCR, False, Customer
                End If
                EstadoImpresora = 20
                
            Else
            
                If colJournal.Count > 0 Then
                    
                    sPasoError = "Imprimo en Journal doc: " & doc
                    EnvioDatosAPos colJournal, False, Journal
                    EstadoImpresora = 10
                    'Consulto estado para validar que fué impreso.
                    EvaluoStatus
                    
                    
                ElseIf colCR.Count > 0 Then
    
                    sPasoError = "Imprimo en cliente doc: " & doc
                    EnvioDatosAPos colCR, False, Customer
                    
                    EstadoImpresora = 20
                    'Consulto estado para validar que fué impreso.
                    EvaluoStatus
    
                End If
            End If
            Exit Sub
        Else
            MsgBox "Para el documento " & doc & " no se obtuvieron filas de impresión.", vbInformation, "ATENCIÓN"
            colDocumentosAImprimir.Remove 1
        End If
    Else
        colDocumentosAImprimir.Remove 1
    End If
    DoEvents
    
    sPasoError = "Sigo al paso siguiente"
    ImprimoDocumentos
    
Exit Sub
errID:
    clsGeneral.OcurrioError "Error al imprimir el ticket. " & vbCrLf & sPasoError, Err.Description

End Sub

Sub ImprimoDocumentos()
Dim bHay As Boolean

    bHay = Not (colDocumentosAImprimir Is Nothing)
    If bHay Then bHay = (colDocumentosAImprimir.Count > 0)
    
    If bHay Then
        'Controlo que la impresora esté en condiciones de imprimir.
        EvaluoStatus
        Exit Sub
    Else
        'Prendo nuevamente el timer de lectura de documentos.
        RegresarAEscuchar
    End If
    Exit Sub
    
End Sub

Private Sub EnviarAArchivoTest(ByVal sDato As String)
Exit Sub

Dim i As Integer
    i = FreeFile
    Open App.Path & "\testingPOS.txt" For Append As #i
    Write #i, sDato
    Close #i

End Sub

Sub EnvioDatosAPos(ByVal colLineas As Collection, ByVal CortarPapel As Boolean, ByVal Recipiente As PrinterIndex)
    Dim bSetFontDefault As Boolean
    Dim iCont As Integer
    Dim oSalida As clsLineaSalida
    bSetFontDefault = True
    If Not mscPOS.PortOpen Then mscPOS.PortOpen = True
    Dim bBarCode As Boolean
    
    mscPOS.Output = oPos.SetPrinter(Recipiente)
    
    For iCont = 1 To colLineas.Count
        Set oSalida = colLineas(iCont)
        bBarCode = False
        
        If Trim(LCase(oSalida.TextoSalida)) = "cutpaper" Then
            mscPOS.Output = oPos.key_CutPaper '& oPos.key_LF
        Else
            If Not (oSalida.Fuente Is Nothing) Then
                bSetFontDefault = True
                Select Case UCase(oSalida.Fuente.TipoFuente)
                    Case "A"
                        If oSalida.Fuente.High And oSalida.Fuente.Wide Then
                            mscPOS.Output = oPos.key_SetFontADoubleWideDoubleHigh
                        ElseIf oSalida.Fuente.High Then
                            mscPOS.Output = oPos.key_SetFontADoubleHigh
                        ElseIf oSalida.Fuente.Wide Then
                            mscPOS.Output = oPos.key_SetFontADoubleWide
                        Else
                            mscPOS.Output = oPos.key_SetFontA
                        End If
                        
                        If oSalida.Fuente.Size > 0 Then
                            Select Case oSalida.Fuente.Size
                                Case 1: mscPOS.Output = oPos.key_SetFontSize12
                                Case 2: mscPOS.Output = oPos.key_SetFontSize15
                                Case 3: mscPOS.Output = oPos.key_SetFontSize17
                            End Select
                        End If
                        
                        If oSalida.Fuente.Invert Then mscPOS.Output = oPos.Key_TextInvertInit
                    
                    Case "B"
                        If oSalida.Fuente.High And oSalida.Fuente.Wide Then
                            mscPOS.Output = oPos.key_SetFontBDoubleWideDoubleHigh
                        ElseIf oSalida.Fuente.High Then
                            mscPOS.Output = oPos.key_SetFontBDoubleHigh
                        ElseIf oSalida.Fuente.Wide Then
                            mscPOS.Output = oPos.key_SetFontBDoubleWide
                        Else
                            mscPOS.Output = oPos.key_SetFontB
                        End If
                        If oSalida.Fuente.Size > 0 Then
                            Select Case oSalida.Fuente.Size
                                Case 1: mscPOS.Output = oPos.key_SetFontSize12
                                Case 2: mscPOS.Output = oPos.key_SetFontSize15
                                Case 3: mscPOS.Output = oPos.key_SetFontSize17
                            End Select
                        End If
                        
                        If oSalida.Fuente.Invert Then mscPOS.Output = oPos.Key_TextInvertInit
                        
                    Case "17"
                        mscPOS.Output = oPos.key_SetFontSize17
                    
                    Case "15"
                        mscPOS.Output = oPos.key_SetFontSize15
                        
                    Case "12"
                        mscPOS.Output = oPos.key_SetFontSize12
                    
                    Case "UPCA"
                        If oSalida.Fuente.Size <> 0 Then mscPOS.Output = oPos.key_SetearHeighCodeBar(oSalida.Fuente.Size)
                        mscPOS.Output = oPos.key_InitBARCODE & oPos.key_FONTCB_UPCA & oSalida.TextoSalida & oPos.key_EndBARCODE
                        bBarCode = True
                    Case "UPCE"
                        If oSalida.Fuente.Size <> 0 Then mscPOS.Output = oPos.key_SetearHeighCodeBar(oSalida.Fuente.Size)
                        mscPOS.Output = oPos.key_InitBARCODE & oPos.key_FONTCB_UPCE & oSalida.TextoSalida & oPos.key_EndBARCODE
                        bBarCode = True
                    Case "JAN13"
                        If oSalida.Fuente.Size <> 0 Then mscPOS.Output = oPos.key_SetearHeighCodeBar(oSalida.Fuente.Size)
                        mscPOS.Output = oPos.key_InitBARCODE & oPos.key_FONTCB_JAN13 & oSalida.TextoSalida & oPos.key_EndBARCODE
                        bBarCode = True
                    Case "JAN8"
                        If oSalida.Fuente.Size <> 0 Then mscPOS.Output = oPos.key_SetearHeighCodeBar(oSalida.Fuente.Size)
                        mscPOS.Output = oPos.key_InitBARCODE & oPos.key_FONTCB_JAN8 & oSalida.TextoSalida & oPos.key_EndBARCODE
                        bBarCode = True
                    
                    Case "CODE39"
                        mscPOS.Output = oPos.key_AlignCenter
                        EnviarAArchivoTest oSalida.TextoSalida & vbTab & IIf(Recipiente = Journal, "DIARIO", "JOURNAL") & " Tamaño = " & oSalida.Fuente.Size
                        If oSalida.Fuente.Size <> 0 Then mscPOS.Output = oPos.key_SetearHeighCodeBar(oSalida.Fuente.Size)
                        mscPOS.Output = oPos.key_SetearWidthCodeBar(2)
                        fnc_Espera
                        mscPOS.Output = oPos.key_InitBARCODE & oPos.key_FONTCB_CODE39 & Trim(oSalida.TextoSalida) & oPos.key_EndBARCODE
                        bBarCode = True
                        mscPOS.Output = oPos.key_AlignLeft
                        
                        
                    Case "ITF"
                        If oSalida.Fuente.Size <> 0 Then mscPOS.Output = oPos.key_SetearHeighCodeBar(oSalida.Fuente.Size)
                        mscPOS.Output = oPos.key_InitBARCODE & oPos.key_FONTCB_ITF & oSalida.TextoSalida & oPos.key_EndBARCODE
                        bBarCode = True
                    Case "CODABAR"
                        If oSalida.Fuente.Size <> 0 Then mscPOS.Output = oPos.key_SetearHeighCodeBar(oSalida.Fuente.Size)
                        mscPOS.Output = oPos.key_InitBARCODE & oPos.key_FONTCB_CODABAR & oSalida.TextoSalida & oPos.key_EndBARCODE
                        bBarCode = True
                    Case "CODE128"
                        If oSalida.Fuente.Size <> 0 Then mscPOS.Output = oPos.key_SetearHeighCodeBar(oSalida.Fuente.Size)
                        mscPOS.Output = oPos.key_InitBARCODE & oPos.key_FONTCB_CODE128 & oSalida.TextoSalida & oPos.key_EndBARCODE
                        bBarCode = True
                        
                    Case "CODE93"
                        If oSalida.Fuente.Size <> 0 Then mscPOS.Output = oPos.key_SetearHeighCodeBar(oSalida.Fuente.Size)
                        mscPOS.Output = oPos.key_InitBARCODE & oPos.key_FONTCB_CODE93 & oSalida.TextoSalida & oPos.key_EndBARCODE
                        bBarCode = True
                        
                    Case Else
                        mscPOS.Output = oPos.key_SetFontA: bSetFontDefault = False
                    
                End Select
            Else
                If bSetFontDefault Then mscPOS.Output = oPos.key_SetFontA: bSetFontDefault = False
            End If
            
            'Alineación izquierda.
            mscPOS.Output = oPos.key_AlignLeft
            
            If InStr(1, oSalida.TextoSalida, "[logo]", vbTextCompare) > 0 Then
                Dim nroLogo As Integer
                If IsNumeric(Replace(oSalida.TextoSalida, "[logo]", "", , , vbTextCompare)) Then
                    'Centro la imagen.
                    EnviarAImpresora oPos.key_AlignCenter
                    EnviarAImpresora oPos.ImprimirLogoAlmacenado(Val(Replace(oSalida.TextoSalida, "[logo]", "", , , vbTextCompare)), 0) ' & Chr(&HD)
                    EnviarAImpresora oPos.key_AlignLeft
                End If
            ElseIf Not bBarCode Then
'                If (InStr(1, oSalida.TextoSalida, "[Key_InitTab]", vbTextCompare) > 0) Then
'                    oSalida.TextoSalida = Replace(oSalida.TextoSalida, "[Key_InitTab]", "", , , vbTextCompare)
'                    Dim vTab() As String
'                    vTab = Split(oSalida.TextoSalida)
'                    Dim iCont As Integer
'                    For iCont = 0 To UBound(vTab)
'                        If IsNumeric(vTab(iCont)) Then
'
'                        End If
'                    Next
'                Else
'                    mscPOS.Output = oSalida.TextoSalida & oPos.key_LF
'                End If

                oSalida.TextoSalida = Replace(oSalida.TextoSalida, "[HighQ]", oPos.key_SetHighQualityOn, , , vbTextCompare)
                oSalida.TextoSalida = Replace(oSalida.TextoSalida, "[/HighQ]", oPos.key_SetHighQualityOFF, , , vbTextCompare)

                oSalida.TextoSalida = Replace(oSalida.TextoSalida, "[/HighQ]", oPos.key_SetHighQualityOFF, , , vbTextCompare)

                oSalida.TextoSalida = Replace(oSalida.TextoSalida, "[Bold]", oPos.Key_TextEnphasizedON, , , vbTextCompare)
                oSalida.TextoSalida = Replace(oSalida.TextoSalida, "[/Bold]", oPos.Key_TextEnphasizedOFF, , , vbTextCompare)
                
                oSalida.TextoSalida = Replace(oSalida.TextoSalida, "[key_LF]", oPos.key_LF, , , vbTextCompare)
                
                oSalida.TextoSalida = Replace(oSalida.TextoSalida, "[key_AlignC]", oPos.key_AlignCenter, , , vbTextCompare)
                
                mscPOS.Output = Replace(oSalida.TextoSalida, "[key_AlignR]", oPos.key_AlignRight, , , vbTextCompare) & oPos.key_LF
                '[Key_InitTab]' + CHAR(160) + '[Key_InitTab]'
                
            End If
            If Not oSalida.Fuente Is Nothing Then
                If oSalida.Fuente.Invert Then mscPOS.Output = oPos.Key_TextInvertEnd
            End If
        End If
    Next
    
    If CortarPapel Then mscPOS.Output = oPos.key_LF & oPos.key_CutPaper
    

End Sub

'Sub InicializoImprimoTabulacion(ByVal Texto As String, ByVal Tabulacion As String)
'Dim vTab() As String
'Dim sTabulacion As String
'
'    vTab = Split(Tabulacion, ";")
'
'    Dim iQ As Byte
'    For iQ = 0 To UBound(vTab)
'
'    Next
'    EnviarAImpresora
'End Sub

Private Sub fnc_Espera()
Dim iCorrer As Long
Dim sPregunta As String
    iCorrer = 0
    Do While iCorrer < 600000
        If sPregunta = "A" Then sPregunta = "B" Else sPregunta = "A"
        If sPregunta = "A" Then sPregunta = "B" Else sPregunta = "A"
        iCorrer = iCorrer + 1
    Loop
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = True
        'Me.WindowState = vbMinimized
        Me.Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmLectura.Enabled = False
    tmStatus.Enabled = False
    If (cboPos.ListIndex >= 0) Then SaveSetting App.Title, "Configuración", "Tickeadora", cboPos.Text
    Unload frmInicio
End Sub



Private Sub lblRollo_DblClick()
    If butActivarPos.Enabled And Val(butActivarPos.Tag) = 0 Then
        ValidarNroRollo
    End If
End Sub

Private Sub MnuArcEstado_Click()
    
    EstadoImpresora = 5
    EvaluoStatus

End Sub

Private Sub MnuArcGrabarLinea_Click()
    EnviarAImpresora oPos.AlmacenarLogo(2, 68, 1, modLogoFuente3.LineaRecta())
End Sub

Private Sub MnuArcGrabarlogo_Click()
 
    Dim sLogo As String
    sLogo = modLogoChico.Horizontal01 & modLogoChico.Horizontal02 & modLogoChico.Horizontal03 & modLogoChico.Horizontal04
    EnviarAImpresora oPos.AlmacenarLogo(1, 40, 10, sLogo)
    'EnviarAImpresora oPos.AlmacenarLogo(1, 53, 11, LogoFuente3)
        
    'EnviarAImpresora oPos.Key_Continue & LogoFuente3_3
End Sub

Private Function DameNroPuerto() As Integer
On Error GoTo errDP
    DameNroPuerto = mscPOS.CommPort
    Exit Function
errDP:
    DameNroPuerto = 1
End Function

Private Sub MnuArchCnfgPuerto_Click()
On Error GoTo errSeteo
Dim iPort As Byte
    
    If Val(butActivarPos.Tag) = 1 Then
        MsgBox "Debe desactivar el pos para realizar cambios.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    iPort = DameNroPuerto
    Dim sResult As String
    sResult = InputBox("Indique el número del puerto a utilizar.", "Puerto de impresora", CStr(iPort))
    
    If IsNumeric(sResult) And Val(sResult) <> iPort Then
        SaveSetting App.Title, "Configuración", "Puerto", sResult
        If mscPOS.PortOpen Then mscPOS.PortOpen = False
        DoEvents
        mscPOS.CommPort = Val(sResult)
        mscPOS.PortOpen = True
    End If
    
    Exit Sub
errSeteo:
    clsGeneral.OcurrioError "Error al setear el puerto, reintente.", Err.Description, "Error"
End Sub

Private Sub MnuArchExit_Click()
    
    If Val(butActivarPos.Tag) = 1 Then
        MsgBox "Primero debe desactivar las lecturas.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    If MsgBox("¿Confirma cerrar el servidor?", vbQuestion + vbYesNo + vbDefaultButton2, "Cerrar") = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub MnuArcImprimirLinea_Click()
    EnviarAImpresora oPos.ImprimirLogoAlmacenado(2, 0) & Chr(&HD)
End Sub

Private Sub MnuArcImprimirLogo_Click()
    'EnviarAImpresora oPos.ImprimirLogoAlmacenado(1, 0)
    EnviarAImpresora oPos.SetPrinter(Customer)
'    EnviarAImpresora oPos.ImprimirLogo(56, 10, LogoFuente, 0)
        'EnviarAImpresora oPos.ImprimirLogo(53, 12, LogoFuente3_1, 0)
        'EnviarAImpresora oPos.ImprimirLogo(64, 1, LineaRecta, 0)
    EnviarAImpresora oPos.ImprimirLogo(40, 10, modLogoChico.Horizontal01 & modLogoChico.Horizontal02 & modLogoChico.Horizontal03 & modLogoChico.Horizontal04, 0)
    
End Sub

Private Sub MnuArcLimpiarMemoria_Click()
    EnviarAImpresora oPos.LimpiarMemoria(logo)
End Sub

Private Sub MnuArcLogoAlmacenado_Click()
    EnviarAImpresora oPos.ImprimirLogoAlmacenado(1, 0) & Chr(&HD)
    'EnviarAImpresora oPos.ImprimirLogoAlmacenado(2, 0) & Chr(&HD)
End Sub

Private Sub MnuFuentes_Click()
    
    EnviarAImpresora oPos.SetPrinter(Customer)
    EnviarAImpresora oPos.Key_TextEnphasizedON
    EnviarAImpresora "Texto size on " & oPos.key_LF
    EnviarAImpresora oPos.Key_TextEnphasizedOFF
    EnviarAImpresora "Texto size off " & oPos.key_LF
    EnviarAImpresora "Texto" & oPos.key_LF
    
    EnviarAImpresora "A izquierda"
    EnviarAImpresora oPos.key_AlignColumnRight
    EnviarAImpresora "derecha"
    EnviarAImpresora oPos.key_LF
    
    EnviarAImpresora "Texto Libre " & oPos.key_LF
    EnviarAImpresora "A izquierda " & oPos.key_AlignColumnRight & " derecha " & oPos.key_LF
    
End Sub

Private Sub MnuHighQualityOFF_Click()
    EnviarAImpresora oPos.key_SetHighQualityOFF
End Sub

Private Sub MnuHighQualityON_Click()
    EnviarAImpresora oPos.key_SetHighQualityOn
End Sub

Private Sub MnuCnfTickeadoraRP_Click()
On Error GoTo errTRP

    If MnuCnfTickeadoraRP.Checked Then
        If MsgBox("Confirma que esta tickeadora no es más la encargada de imprimir las cuotas pagas por Giros?", vbQuestion + vbYesNo, "QUITAR TICKEADORA") = vbNo Then Exit Sub
    Else
        If MsgBox("Confirma asignar la tickeadora como la encargada de imprimir las cuotas pagas por Giros?", vbQuestion + vbYesNo, "ASIGNAR TICKEADORA") = vbNo Then Exit Sub
    End If
    'Si quito y es otro nro. no hago nada.
    'Si asigno si o si sobreescribo.
    If MnuCnfTickeadoraRP.Checked Then
        Cons = "UPDATE Parametro SET ParValor = Null WHERE ParNombre = 'TickeadoraGirosCuotas' AND ParValor = " & cboPos.ItemData(cboPos.ListIndex)
    Else
        Cons = "UPDATE Parametro SET ParValor = " & cboPos.ItemData(cboPos.ListIndex) & " WHERE ParNombre = 'TickeadoraGirosCuotas'"
    End If
    cBase.Execute Cons
    MnuCnfTickeadoraRP.Checked = Not MnuCnfTickeadoraRP.Checked
    Exit Sub
errTRP:
    clsGeneral.OcurrioError "Error en la asignación.", Err.Description, "Cuotas de Giros"
End Sub

Private Sub MnuInformeZ_Click()
On Error GoTo errMIF
    
    'Valido que no queden recibos sin imprimir para la sucursal.
    Dim SDoc As String
    Cons = "SELECT DocSerie, DocNumero " & _
            "FROM cgsa.dbo.Documento INNER JOIN cgsa.dbo.Usuario ON UsuCodigo = DocUsuario " & _
            "WHERE DocTipo IN (5) AND DocFecha BETWEEN '" & Format(Now, "yyyy/mm/dd 00:00:00") & "' AND '" & Format(Now, "yyyy/mm/dd 23:59:59") & "'" & _
            "AND DocCodigo NOT IN (SELECT TAIDocumento FROM cgsa.dbo.TicketsAImprimir)"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        SDoc = RsAux("Docserie") & "-" & RsAux("DocNumero")
        RsAux.Close
        MsgBox "Quedan recibos como el " & SDoc & " que no fue impreso, no puede emitir el informe Z", vbCritical, "ATENCIÓN"
        Exit Sub
    End If
    RsAux.Close
    
    Cons = "SELECT DocSerie, DocNumero " & _
            "FROM cgsa.dbo.Documento INNER JOIN cgsa.dbo.Usuario ON UsuCodigo = DocUsuario " & _
            "WHERE DocTipo IN (5) AND DocFecha BETWEEN '" & Format(Now, "yyyy/mm/dd 00:00:00") & "' AND '" & Format(Now, "yyyy/mm/dd 23:59:59") & "'" & _
            "AND DocCodigo IN (SELECT TAIDocumento FROM cgsa.dbo.TicketsAImprimir WHERE TAIEstado <> 1 or TAIRollo is null)"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        
        SDoc = RsAux("Docserie") & "-" & RsAux("DocNumero")
        RsAux.Close
        MsgBox "Quedan recibos como el " & SDoc & " que no fue impreso, no puede emitir el informe Z", vbCritical, "ATENCIÓN"
        Exit Sub
    End If
    RsAux.Close
    
    If MsgBox("¿Confirma realizar el informe del cierre del día?", vbQuestion + vbYesNo + vbDefaultButton2, "Cerrar el día") = vbYes Then
        ApagarServidor
        ImprimoPrimerDocumento (True)
    End If
    
    Exit Sub
errMIF:
    clsGeneral.OcurrioError "Error al almacenar la información", Err.Description, "Servidor de impresión"
End Sub

Private Sub MnuRePrintAPartirDe_Click()
Dim sDesde As String, sValidar As String

    sDesde = InputBox("Ingrese desde que número de documento desea reimprimir (serie-número)", "Reimpresión de tickets")
    If Trim(sDesde) = "" Then Exit Sub
    
    Dim sHasta As String
    sHasta = InputBox("Ingrese hasta que número desea reimprimir (número, la serie es opcional)", "Reimpresión de tickets")
    If Trim(sHasta) = "" Then Exit Sub Else sHasta = Trim(sHasta)
    
    
    'Valido serie y número del recibo.
    sValidar = InputBox("¿Está seguro que ud desea reimprimir muchos tickets?" & vbCrLf & vbCrLf & _
                "Para confirmar escriba la palabra 'Muchos' " & vbCrLf & "ROGAMOS PRESTE MÁXIMA ATENCIÓN", "Confirmación", "")
    
    If LCase(Trim(sValidar)) = "muchos" Then
        Dim idDoc As Long, idDocHasta As Long
        Dim sSerie As String
        idDoc = BuscarTicketsAImprimir(sDesde, False, sSerie)
        
        If idDoc > 0 Then
            If IsNumeric(Mid(sHasta, 1, 1)) Then sHasta = sSerie & sHasta
            idDocHasta = BuscarTicketsAImprimir(sHasta, False)
        End If
        
        Dim usr As Integer
        Dim defensa As String
        
        usr = PidoSucesoReimpresion(defensa)
        
        If usr = 0 Then MsgBox "Debe indicar un suceso", vbExclamation, "ATENCIÓN": Exit Sub
        
        If idDoc > 0 And idDocHasta > 0 Then
            Cons = "UPDATE TicketsAImprimir SET TAIEstado = Null WHERE TAIEstado = 1 AND TAIDocumento >= " & idDoc & _
                " AND TAIDocumento <= " & idDocHasta & " AND TAITickeadora = " & cboPos.ItemData(cboPos.ListIndex)
            cBase.Execute (Cons)
            
            'Grabo para el primero y para el último.
            GraboSuceso idDoc, sDesde, defensa, usr
            If idDoc <> idDocHasta Then GraboSuceso idDocHasta, sHasta, defensa, usr
            
            MsgBox "TICKETS enviados a reimpresión.", vbInformation, "Estado modificado"
        End If
    End If
    Exit Sub
    
errSave:
    clsGeneral.OcurrioError "Error al modificar el estado de los tickets.", Err.Description, "Error"
    Screen.MousePointer = 0

End Sub

Function BuscarTicketsAImprimir(ByVal nroDocumento As String, ByVal EsCI As Boolean, Optional ByRef Serie As String) As Long
On Error GoTo errTAI

    If Not EsCI Then
        Dim mDSerie As String, mDNumero As Long
        If InStr(nroDocumento, "-") <> 0 Then
            mDSerie = Mid(nroDocumento, 1, InStr(nroDocumento, "-") - 1)
            mDNumero = Val(Mid(nroDocumento, InStr(nroDocumento, "-") + 1))
        Else
            nroDocumento = Replace(nroDocumento, " ", "")
            If IsNumeric(Mid(nroDocumento, 2, 1)) Then
                mDSerie = Mid(nroDocumento, 1, 1)
                mDNumero = Val(Mid(nroDocumento, 2))
            Else
                mDSerie = Mid(nroDocumento, 1, 2)
                mDNumero = Val(Mid(nroDocumento, 3))
            End If
        End If
        
        
        If mDNumero = 0 Then
            MsgBox "Formato de documento incorrecto.", vbInformation, "Atención"
            Exit Function
        End If
    End If

    Dim idDoc As Long
    On Error GoTo errTAI
    Cons = "SELECT DocCodigo, DocFecha Fecha, DocSerie + '-' + CAST(DocNumero as varchar(6)) Documento" & _
        " FROM Documento INNER JOIN TicketsAImprimir ON TAIDocumento = DocCodigo "

    If EsCI Then
        Cons = Cons & " INNER JOIN Cliente ON DocCliente = CliCodigo AND CliCIRUC = '" & Trim(nroDocumento) & _
                        "' WHERE DocAnulado = 0 And DocFecha BETWEEN '" & Format(DateAdd("n", -30, Now), "yyyy/mm/dd hh:nn:ss") & _
                        "' AND '" & Format(Now, "yyyy/mm/dd hh:nn:ss") & "'"
        Dim objAyuda As New clsListadeAyuda
        If objAyuda.ActivarAyuda(cBase, Cons, 6000, 1, "Tickets de un cliente") > 0 Then
            idDoc = objAyuda.RetornoDatoSeleccionado(0)
        End If
        Set objAyuda = Nothing
        BuscarTicketsAImprimir = idDoc
        Exit Function
    Else
        Cons = Cons & " WHERE DocSerie = '" & mDSerie & "' AND DocNumero = " & mDNumero & _
        " AND DocAnulado = 0 And DocFecha > '" & Format(DateAdd("h", -8, Now), "yyyy/mm/dd hh:nn:ss") & "'"
    End If
    
    Dim rsD As rdoResultset
    Set rsD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If rsD.EOF Then
        MsgBox "No se encontró un ticket con esos datos o el mismo fue anulado.", vbExclamation, "ATENCIÓN"
    ElseIf Format(rsD("Fecha"), "dd/mm/yyyy") <> Format(Date, "dd/mm/yyyy") Then
        MsgBox "El ticket no fue impreso hoy, verifique", vbExclamation, "Atención"
    Else
        Serie = mDSerie
        idDoc = rsD("DocCodigo")
    End If
    rsD.Close
    BuscarTicketsAImprimir = idDoc
    Exit Function
errTAI:
    clsGeneral.OcurrioError "Error al buscar el ticket a reimprimir", Err.Description, "Tickets"
End Function

Private Sub MnuRePrintNroDoc_Click()
On Error GoTo errMRP
    Dim SDoc As String
    SDoc = InputBox("Ingrese serie y número del ticket", "Reimprimir un documento", "#-######")
    If SDoc <> "" Then
        Dim idDoc As Long
        idDoc = BuscarTicketsAImprimir(SDoc, False)
        If idDoc > 0 Then
            Dim usr As Integer
            Dim defensa As String
            usr = PidoSucesoReimpresion(defensa)
        
            If usr > 0 Then
                
                Cons = "UPDATE TicketsAImprimir SET TAIEstado = Null, TAITickeadora = " & cboPos.ItemData(cboPos.ListIndex) & " WHERE TAIDocumento = " & idDoc
                cBase.Execute (Cons)
                
                GraboSuceso idDoc, SDoc, defensa, usr
            Else
                MsgBox "Debe ingresar el suceso para continuar.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
            MsgBox "TICKET enviado a reimpresión.", vbInformation, "Estado modificado"
        End If
    End If
    Exit Sub
errMRP:
clsGeneral.OcurrioError "Error al almacenar la información", Err.Description, "Servidor de impresión"
End Sub

Private Sub MnuTicketCI_Click()
On Error GoTo errSave

    Dim sDesde As String
    sDesde = InputBox("Ingrese la cédula del primer ticket a reimprimir (Sin puntos ni guión).", "Reimpresión de tickets")
    sDesde = Replace(sDesde, ".", "")
    sDesde = Replace(sDesde, "-", "")
    If Trim(sDesde) = "" Then Exit Sub
    
    Dim idDoc As Long
    idDoc = BuscarTicketsAImprimir(sDesde, True)

    If idDoc > 0 Then
    
        Dim usr As Integer
        Dim defensa As String
        usr = PidoSucesoReimpresion(defensa)
        
        If usr > 0 Then
    
            Cons = "UPDATE TicketsAImprimir SET TAIEstado = Null WHERE TAIEstado = 1 AND TAIDocumento = " & idDoc & _
                " AND TAITickeadora = " & cboPos.ItemData(cboPos.ListIndex)
            cBase.Execute (Cons)
            GraboSuceso idDoc, sDesde, defensa, usr
        
            MsgBox "TICKET enviado a reimpresión.", vbInformation, "Estado modificado"
        End If
    End If
    Exit Sub
    
errSave:
    clsGeneral.OcurrioError "Error al modificar el estado de los tickets.", Err.Description, "Error"
    Screen.MousePointer = 0
End Sub

Private Sub MnuTicketsRangoCI_Click()
Dim sDesde As String, sValidar As String

    sDesde = InputBox("Ingrese la cédula del primer ticket a reimprimir (Sin puntos ni guión).", "Reimpresión de tickets")
    sDesde = Replace(sDesde, ".", "")
    sDesde = Replace(sDesde, "-", "")
    If Trim(sDesde) = "" Then Exit Sub
    
    Dim sHasta As String
    sHasta = InputBox("Ingrese la cédula del último ticket a reimprimir (sin puntos ni guión).", "Reimpresión de tickets")
    sHasta = Replace(sHasta, ".", "")
    sHasta = Replace(sHasta, "-", "")
    If Trim(sHasta) = "" Then Exit Sub
    
    'Valido serie y número del recibo.
    sValidar = InputBox("¿Está seguro que ud desea reimprimir muchos tickets?" & vbCrLf & vbCrLf & _
                "Para confirmar escriba la palabra 'Muchos' " & vbCrLf & "ROGAMOS PRESTE MÁXIMA ATENCIÓN", "Confirmación", "")
    
    If LCase(Trim(sValidar)) = "muchos" Then
        Dim idDoc As Long, idDocHasta As Long
        Dim sSerie As String
        idDoc = BuscarTicketsAImprimir(sDesde, True)
        
        If idDoc > 0 Then
            idDocHasta = BuscarTicketsAImprimir(sHasta, True)
        End If
        
        If idDoc > 0 And idDocHasta >= idDoc Then
        
            Dim usr As Integer
            Dim defensa As String
            usr = PidoSucesoReimpresion(defensa)
            
            If usr > 0 Then
                Cons = "UPDATE TicketsAImprimir SET TAIEstado = Null WHERE TAIEstado = 1 AND TAIDocumento >= " & idDoc & _
                    " AND TAIDocumento <= " & idDocHasta & " AND TAITickeadora = " & cboPos.ItemData(cboPos.ListIndex)
            
                cBase.Execute (Cons)
                
                GraboSuceso idDoc, "", defensa, usr
                GraboSuceso idDocHasta, "", defensa, usr
            
                MsgBox "TICKETS enviados a reimpresión.", vbInformation, "Estado modificado"
            End If
        End If
    End If
    Exit Sub
errSave:
    clsGeneral.OcurrioError "Error al modificar el estado de los tickets.", Err.Description, "Error"
    Screen.MousePointer = 0
End Sub

Private Sub mscPOS_OnComm()

    Select Case mscPOS.CommEvent
       
       Case comEvReceive
       IndicarActividad "Recibiendo info POS ..."
        If EstadoImpresora = 5 Then
            RespuestaEstado mscPOS.Input
        Else
            InformacionPos mscPOS.Input
        End If
    End Select
    
End Sub

Private Sub tmCliente_Timer()
    
    tmCliente.Enabled = False
    If EstadoImpresora = 10 Then
        Dim i As Long
        i = 0
        Dim sDemoro As String
        Do While i < 409000
            i = i + 1
            If sDemoro <> "" Then
                sDemoro = "A"
            Else
                sDemoro = ""
            End If
        Loop
        EnvioDatosAPos colCR, False, Customer
        EstadoImpresora = 20
        EvaluoStatus
        
    Else
        EstadoImpresora = 0
        If colDocumentosAImprimir.Count > 0 Then
            colDocumentosAImprimir.Remove (1)
            ImprimoDocumentos   'vuelvo a imprimir.
        End If
    End If
    
End Sub

Private Sub tmLectura_Timer()
    
    tmLectura.Enabled = False
    frmInicio.CambiarIcono "Servidor activo", Normal
    butActivarPos.Enabled = False
    If BuscoDocumentosAImprimir Then ImprimoDocumentos Else tmLectura.Enabled = True
    butActivarPos.Enabled = True
'    tmLectura.Enabled = True
    
End Sub

Sub RegresarAEscuchar()
    
    'Pasaron 10 segundos y no tengo respuesta de la impresora.
    tmStatus.Enabled = False
    'Reinicio todo de nuevo.
    
    butActivarPos.Enabled = True
    Set colDocumentosAImprimir = Nothing
    tmLectura.Enabled = True
    iStatus = 0
    
End Sub

Private Sub tmStatus_Timer()
    
    If Val(butActivarPos.Tag) = 0 And EstadoImpresora = 0 Then ApagarServidor: Exit Sub
    
    'Imprimo el primer elemento en la collección.
    If iStatus = 10 Then
    
        If DateDiff("s", CDate(tmStatus.Tag), Now) > 6 Then
            
            frmInicio.CambiarIcono "Servidor en error", Error
            Me.Visible = True
            tmStatus.Enabled = False
            
            If EstadoImpresora = 10 Or EstadoImpresora = 20 Then
                MsgBox "ATENCIÓN no se obtiene respuesta de la impresora." & vbCrLf & vbCrLf & _
                    "El último documento no fué impreso en su totalidad por favor valide el mismo.", vbExclamation, "ATENCIÓN"
            Else
                MsgBox "No se obtiene respuesta de la impresora, verifiqué que este encendida y con las bandejas cerradas." & _
                    vbCrLf & vbCrLf & "último buffer: " & sBufferMsgPrint _
                    , vbCritical, "ATENCIÓN"
            
            End If
            
            EstadoImpresora = 0
            Set colDocumentosAImprimir = Nothing
            butActivarPos_Click
            Exit Sub
            
        End If
        Exit Sub
    End If
    tmStatus.Enabled = False
        
End Sub

Public Sub EvaluoStatus()
    
'    If Val(butActivarPos.Tag) = 0 And Not SolicitoEstado Then ApagarServidor: Exit Sub
    
    'Testeo el status de la impresora
    If Not mscPOS.PortOpen Then mscPOS.PortOpen = True
    
    IndicarActividad "Consultando estado ..."
        
    iStatus = 10        'Indico que estoy escuchando a la impresora.
    sBufferMsgPrint = ""
    
    'Pido el status limpio el buffer.
    tmStatus.Enabled = True
    tmStatus.Interval = 400
    tmStatus.Tag = Now
    
    EnviarAImpresora oPos.Status
        
End Sub

Private Sub InformacionPos(ByVal msg As String)
Dim sPaso As String

    If Val(butActivarPos.Tag) = 0 Then ApagarServidor: Exit Sub
    
    sBufferMsgPrint = sBufferMsgPrint & msg
    IndicarActividad "Recibiendo " & Len(sBufferMsgPrint) & " bytes "
    
    If Len(sBufferMsgPrint) = 10 Then
        
        If Asc(Mid(sBufferMsgPrint, 2, 1)) = 10 Then
            'Respuesta a status.
            iStatus = 1
            
            'Datos conocidos pos 3 = 72 --> tengo la bandeja cliente sin papel o abierta.
            '               pos 7 = 96 -> journal abierta
            '                     = 32 --> ok,
            '                     = 160 --> journal abierta.
            '                     = 224 --> sin papel y abierta.
                        
            Dim pos3 As Integer
            Dim pos7 As Integer
            pos3 = Asc(Mid(sBufferMsgPrint, 3, 1))
            pos7 = Asc(Mid(sBufferMsgPrint, 7, 1))
            
            If (pos3 = 9 And pos7 = 32) And EstadoImpresora = 10 Or EstadoImpresora = 20 Then
                
                sBufferMsgPrint = ""
                tmCliente.Enabled = True
                
                
            ElseIf (pos3 = 8 Or pos3 = 137) And pos7 = 32 Then
                
                sBufferMsgPrint = ""
                tmStatus.Enabled = False
                tmStatus.Tag = Now
                
                If pos3 = 9 Then
                'Esta ocupada por lo tanto envío de nuevo para hacer espera.
                    IndicarActividad "POS Ocupada"
                    ImprimoDocumentos
                Else
                'OK se puede imprimir.
                    ImprimoPrimerDocumento
                End If
                
            Else
            
                Dim sBytes As String
                Dim iQ As Byte
                For iQ = 1 To Len(sBufferMsgPrint)
                    sBytes = sBytes & ", " & Asc(Mid(sBufferMsgPrint, iQ, 1))
                Next

                'No se puede imprimir.      & vbCrLf & sBytes
                sBufferMsgPrint = ""
                
               
                'RegresarAEscuchar
                'Cambié para apagar de forma que solicite nuevamente el rollo.
                Set colDocumentosAImprimir = Nothing
                ApagarServidor
                frmInicio.CambiarIcono "Pausado por error", Error
                Me.Visible = True
                
                MsgBox "La impresora indica que no se puede imprimir el documento." & vbCrLf & sBytes, vbExclamation, "ATENCIÓN"
                
                
                Exit Sub

            End If

        ElseIf Len(sBufferMsgPrint) > 8 Then
        
            sBufferMsgPrint = ""
            
        End If
                
    End If
    
End Sub

Private Sub RespuestaEstado(ByVal msg As String)
Dim sPaso As String

    
    sBufferMsgPrint = sBufferMsgPrint & msg
    IndicarActividad "Recibiendo " & Len(sBufferMsgPrint) & " bytes "
    
    If Len(sBufferMsgPrint) = 10 Then
        
        If Asc(Mid(sBufferMsgPrint, 2, 1)) = 10 Then
            'Respuesta a status.
            iStatus = 1
            
            'Datos conocidos pos 3 = 72 --> tengo la bandeja cliente sin papel o abierta.
            '               pos 7 = 96 -> journal abierta
            '                     = 32 --> ok,
            '                     = 160 --> journal abierta.
            '                     = 224 --> sin papel y abierta.
                        
            Dim pos3 As Integer
            Dim pos7 As Integer
            pos3 = Asc(Mid(sBufferMsgPrint, 3, 1))
            pos7 = Asc(Mid(sBufferMsgPrint, 7, 1))
            
            If (pos3 = 8 Or pos3 = 9) And pos7 = 32 Then
                sBufferMsgPrint = ""
                tmStatus.Enabled = False
                tmStatus.Tag = Now
                
                If pos3 = 9 Then
                'Esta ocupada por lo tanto envío de nuevo para hacer espera.
                    IndicarActividad "POS Ocupada"
                    MsgBox "Impresora ocupada", vbExclamation, "Atención"
                Else
                'OK se puede imprimir.
                    MsgBox "Impresora lista.", vbInformation, "Estado correcto"
                End If
                
            Else
            
                MsgBox "Estado incorrecto para imprimir.", vbExclamation, "Atención"

            End If
            EstadoImpresora = 0

        ElseIf Len(sBufferMsgPrint) > 8 Then
        
            sBufferMsgPrint = ""
            
            EstadoImpresora = 0
        End If
                
    End If
    
End Sub

Private Sub CargoCombo(Consulta As String, combo As Control, Optional Seleccionado As String = "")
Dim RsAuxiliar As rdoResultset
Dim iSel As Integer: iSel = -1     'Guardo el indice del seleccionado
    
On Error GoTo ErrCC
    
    Screen.MousePointer = 11
    combo.Clear
    Set RsAuxiliar = cBase.OpenResultset(Consulta, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAuxiliar.EOF
        combo.AddItem Trim(RsAuxiliar(1))
        combo.ItemData(combo.NewIndex) = RsAuxiliar(0)
        
        If Trim(RsAuxiliar(1)) = Trim(Seleccionado) Then iSel = combo.ListCount
        RsAuxiliar.MoveNext
    Loop
    RsAuxiliar.Close
    
    If iSel = -1 Then combo.ListIndex = iSel Else combo.ListIndex = iSel - 1
    Screen.MousePointer = 0
    Exit Sub
    
ErrCC:
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al cargar el combo: " & Trim(combo.Name) & "." & vbCrLf & Err.Description, vbCritical, "ERROR"
End Sub

Private Function PidoSucesoReimpresion(ByRef defensa As String) As Integer
On Error GoTo errRSR
Dim user As Integer
Dim defensaTxt As String
    Dim objSuceso As New clsSuceso
    objSuceso.ActivoFormulario paCodigoDeUsuario, "Reimpresión de Documentos", cBase
    user = objSuceso.RetornoValor(usuario:=True)
    defensaTxt = objSuceso.RetornoValor(defensa:=True)
    Set objSuceso = Nothing
    Me.Refresh
    If user = 0 Then Screen.MousePointer = 0: Exit Function 'Abortó el ingreso del suceso
    PidoSucesoReimpresion = user
    defensa = defensaTxt
    '---------------------------------------------------------------------------------------------
    Exit Function
errRSR:
    clsGeneral.OcurrioError "Error al intentar registrar el suceso.", Err.Description, "Error"
End Function

Private Sub GraboSuceso(ByVal documento As Long, ByVal serienumero As String, ByVal defensa As String, ByVal usuario As Integer)
Dim fecha As Date

    '10 = TipoSuceso.Reimpresiones
    clsGeneral.RegistroSuceso cBase, Now, 10, paCodigoDeTerminal, usuario, documento, _
                               Descripcion:=serienumero, defensa:=Trim(defensa)

End Sub

Private Function EmitirCFE(ByVal idDocumento As Long) As Boolean
On Error GoTo errE
'    With New clsCGSAEFactura
'        Set .Connect = cBase
'        Dim xmlRsp As String
'        EmitirCFE = .GenerarEComprobante(idDocumento)
'    End With
    Exit Function
errE:
End Function

