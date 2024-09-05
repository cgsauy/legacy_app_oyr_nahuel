VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form DeMovimientoStock 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalle del Movimiento"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDetMovStock.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBotones 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   6495
      TabIndex        =   14
      Top             =   2880
      Width           =   6495
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   480
         Picture         =   "frmDetMovStock.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   120
         Picture         =   "frmDetMovStock.frx":0744
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   960
         Picture         =   "frmDetMovStock.frx":0A86
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   0
         Width           =   310
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Información General"
      ForeColor       =   &H000000C0&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Terminal:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lTerminal 
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   19
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lArticulo 
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   960
         Width           =   5415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Artículo:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lDocumento 
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   720
         Width           =   5535
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Documento:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lValor 
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   1200
         Width           =   5415
      End
      Begin VB.Label lSucursal 
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4560
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Local:"
         Height          =   255
         Left            =   3840
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   3840
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha/Hora:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lFecha 
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4560
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   3315
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "[Ctrl+A] Anterior"
            TextSave        =   "[Ctrl+A] Anterior"
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Text            =   "[Ctrl+S] Siguiente"
            TextSave        =   "[Ctrl+S] Siguiente"
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Text            =   "[Ctrl+X] Salir"
            TextSave        =   "[Ctrl+X] Salir"
            Key             =   "msg"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4366
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lTexto 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   6615
   End
End
Attribute VB_Name = "DeMovimientoStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private aTexto As String
Private gLista As vsFlexGrid
Private iTipoMov As Integer
Private aComentario As String

Public Property Get pTipoMovimiento() As Integer
    pTipoMovimiento = iTipoMov
End Property
Public Property Let pTipoMovimiento(iTipo As Integer)
    iTipoMov = iTipo
End Property
Public Property Get pLista() As vsFlexGrid
End Property
Public Property Let pLista(nLista As vsFlexGrid)
    Set gLista = nLista
End Property

Private Sub bAnterior_Click()

    On Error GoTo errSuceso
    
    If gLista.Row > 1 Then
        Screen.MousePointer = 11
        gLista.Row = gLista.Row - 1
        CargoDatosMovimiento CLng(gLista.Cell(flexcpData, gLista.Row, 0))
        Screen.MousePointer = 0
        bSiguiente.Enabled = True
    Else
        bAnterior.Enabled = False
    End If
    Exit Sub

errSuceso:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el movimiento anterior.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bSiguiente_Click()

    On Error GoTo errSuceso
    If gLista.Row < gLista.Rows - 1 Then
        Screen.MousePointer = 11
        gLista.Row = gLista.Row + 1
        CargoDatosMovimiento CLng(gLista.Cell(flexcpData, gLista.Row, 0))
        Screen.MousePointer = 0
        bAnterior.Enabled = True
    Else
        bSiguiente.Enabled = False
    End If
    Exit Sub

errSuceso:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el movimiento siguiente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyA: If bAnterior.Enabled Then bAnterior_Click
            Case vbKeyS: If bSiguiente.Enabled Then bSiguiente_Click
            Case vbKeyX: Unload Me
        End Select
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    CargoDatosMovimiento CLng(gLista.Cell(flexcpData, gLista.Row, 0))
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el movimiento.", Err.Description
End Sub

Private Sub CargoDatosMovimiento(Codigo As Long)
Dim aDocumento As Long
Dim aTipoDoc As Long
Dim aArticulo As Long

    aDocumento = 0: aArticulo = 0
    LimpioFicha
    
    If iTipoMov = TipoEstadoMercaderia.Fisico Then
    
        Cons = "Select * from MovimientoStockFisico" _
                    & " Left Outer Join Terminal On MSFTerminal = TerCodigo" _
                & " , Usuario, Local Where MSFCodigo = " & Codigo _
                & " And MSFUsuario = UsuCodigo And MSFLocal = LocCodigo"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        
        If Not RsAux.EOF Then
            lFecha.Caption = Format(RsAux!MSFFecha, "Ddd d-Mmm yyyy hh:mm:ss")
            lFecha.Tag = RsAux!MSFFecha
            lUsuario.Caption = Trim(RsAux!UsuApellido1) & ", " & Trim(RsAux!UsuNombre1)
            lSucursal.Caption = Trim(RsAux!LocNombre)
            If Not IsNull(RsAux!TerNombre) Then lTerminal.Caption = Trim(RsAux!TerNombre)
            lValor.Caption = RsAux!MSFCantidad
            
            If Not IsNull(RsAux!MSFDocumento) Then
                aDocumento = RsAux!MSFDocumento
                aTipoDoc = RsAux!MSFTipoDocumento
            End If
            
            aArticulo = RsAux!MSFArticulo
            
        End If
        
        RsAux.Close
    Else
    
        Cons = "Select * from MovimientoStockEstado, Usuario, Local" _
                & " Where MSECodigo = " & Codigo _
                & " And MSEUsuario = UsuCodigo" _
                & " And MSELocal = LocCodigo"
                
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        
        If Not RsAux.EOF Then
            lFecha.Caption = Format(RsAux!MSEFecha, "Ddd d-Mmm yyyy hh:mm:ss")
            lUsuario.Caption = Trim(RsAux!UsuApellido1) & ", " & Trim(RsAux!UsuNombre1)
            lSucursal.Caption = Trim(RsAux!LocNombre)
    
            lValor.Caption = RsAux!MSECantidad
            
            If Not IsNull(RsAux!MSEDocumento) Then
                aDocumento = RsAux!MSEDocumento
                aTipoDoc = RsAux!MSETipoDocumento
            End If
            aArticulo = RsAux!MSEArticulo
        End If
        
        RsAux.Close
    End If
    
    If aDocumento <> 0 Then CargoDatosDocumento aDocumento, aTipoDoc, aArticulo
    If aArticulo <> 0 Then CargoDatosArticulo aArticulo
    
End Sub

Private Sub CargoDatosDocumento(Codigo As Long, Tipo As Long, Optional IdArticulo As Long = 0)
    On Error GoTo errDoc
    aTexto = ""
    aComentario = ""
    
    Select Case Tipo
        'Ventas Contado/Credito A Domicilio-----------------------------------------------------------------------
        Case TipoDocumento.ContadoDomicilio, TipoDocumento.CreditoDomicilio
            Cons = "Select * from VentaTelefonica, Sucursal" _
                    & " Where VTeCodigo =  " & Codigo _
                    & " And VTeSucursal = SucCodigo"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            
            aTexto = Trim(RsAux!SucAbreviacion) & " "
            
            Select Case RsAux!VTeTipo
                Case TipoDocumento.ContadoDomicilio: aTexto = aTexto & "Venta Telefónica Contado "
                Case TipoDocumento.CreditoDomicilio: aTexto = aTexto & "Venta Telefónica Crédito "
            End Select
            
            aTexto = aTexto & " N°: " & Trim(RsAux!VTeCodigo) & " " _
                          & "(" & Format(RsAux!VTeFechaLlamado, "d/mm/yy hh:mm") & ")"
                
            RsAux.Close
        '------------------------------------------------------------------------------------------------------------------
    
        'Documentos de Compra de Mercaderia----------------------------------------------------------------------
        Case TipoDocumento.CompraCarta, TipoDocumento.Compracontado, TipoDocumento.CompraCredito, _
                TipoDocumento.CompraRemito, TipoDocumento.CompraCarpeta
    
            Cons = "Select * from RemitoCompra, Sucursal" _
                    & " Where RCoCodigo =  " & Codigo _
                    & " And RCoLocal = SucCodigo"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            aTexto = "Compra de Mercadería "
            If Not RsAux.EOF Then
                
                aComentario = " Sucursal: " & Trim(RsAux!SucAbreviacion) & Chr(vbKeyReturn) _
                                   & "  Documento: " & RetornoNombreDocumento(RsAux!RCoTipo) & " "
                
                If Not IsNull(RsAux!RCoSerie) Then aComentario = aComentario & Trim(RsAux!RCoSerie)
                aComentario = aComentario & " " & RsAux!RCoNumero & " " _
                              & "(" & Format(RsAux!RCoFecha, "d/mm/yy hh:mm") & ")"
                    
            Else
                aTexto = aTexto & " " & RetornoNombreDocumento(CInt(Tipo)) & " "
            End If
            RsAux.Close
            
        Case TipoDocumento.CompraNotaCredito, TipoDocumento.CompraNotaDevolucion
            Cons = "Select * from Compra, ProveedorCliente" _
                    & " Where ComCodigo =  " & Codigo _
                    & " And ComProveedor = PClCodigo"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            aTexto = "Compra de Mercadería "
            If Not RsAux.EOF Then
                
                aComentario = "Documento: " & RetornoNombreDocumento(RsAux!ComTipoDocumento) & " "
                
                If Not IsNull(RsAux!ComSerie) Then aComentario = aComentario & Trim(RsAux!ComSerie)
                aComentario = aComentario & " " & RsAux!ComNumero & " " _
                              & "(" & Format(RsAux!ComFecha, "d/mm/yy hh:mm") & ")" & Chr(vbKeyReturn)
                If Not IsNull(RsAux!PClNombre) Then aComentario = aComentario & Trim(RsAux!PClNombre)
                    
            Else
                aTexto = aTexto & " " & RetornoNombreDocumento(CInt(Tipo)) & " "
            End If
            RsAux.Close
            
        '------------------------------------------------------------------------------------------------------------------
        
        Case TipoDocumento.Traslados
            Cons = "Select Traspaso.*, LocalO.LocNombre Origen, LocalD.LocNombre Destino, LocalI.LocNombre Intermediario  from Traspaso, Local LocalO, Local LocalD, Local LocalI " _
                   & " Where TraCodigo = " & Codigo _
                   & " And TraLocalOrigen *= LocalO.LocCodigo " _
                   & " And TraLocalDestino *= LocalD.LocCodigo " _
                   & " And TraLocalIntermedio *= LocalI.LocCodigo "
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux.EOF Then
                aTexto = "Traslado de Mercadería Nº " & Codigo & " (" & Format(RsAux!TraFecha, "Ddd d/Mmm/yyyy hh:mm") & ")"
                
                'Origen ---> Destino----------------------
                aComentario = " Origen: "
                If Not IsNull(RsAux!TraLocalOrigen) Then aComentario = aComentario & Trim(RsAux!Origen) Else aComentario = aComentario & "N/D"
                aComentario = aComentario & " --> Destino: "
                If Not IsNull(RsAux!TraLocalDestino) Then aComentario = aComentario & Trim(RsAux!Destino) Else aComentario = aComentario & "N/D"
                If Not IsNull(RsAux!TraLocalIntermedio) Then aComentario = aComentario & "        Intermediario: " & Trim(RsAux!Intermediario)
                aComentario = aComentario & Chr(vbKeyReturn) & "  "
                
                aComentario = aComentario & "Realizado por: " & BuscoUsuario(RsAux!TraUsuarioInicial, True)
                If Not IsNull(RsAux!TraUsuarioFinal) Then aComentario = aComentario & "         Recibió: " & BuscoUsuario(RsAux!TraUsuarioFinal, True)
                If Not IsNull(RsAux!TraUsuarioReceptor) Then aComentario = aComentario & "       Controló: " & BuscoUsuario(RsAux!TraUsuarioReceptor, True)
                aComentario = aComentario & Chr(vbKeyReturn) & "  "
                
                If Not IsNull(RsAux!TraFechaEntregado) Then aComentario = aComentario & "Entregado: " & Format(RsAux!TraFechaEntregado, "d/Mmm/yyyy hh:mm") Else aComentario = aComentario & " Entregado: NO"
                If Not IsNull(RsAux!TraFImpreso) Then aComentario = aComentario & "                         Impreso: " & Format(RsAux!TraFImpreso, "d/Mmm/yyyy hh:mm") Else aComentario = aComentario & " Impreso: NO"
                
                aComentario = aComentario & Chr(vbKeyReturn) & "  "
                If Not IsNull(RsAux!TraComentario) Then aComentario = aComentario & Trim(RsAux!TraComentario)
                
            End If
            RsAux.Close
            '------------------------------------------------------------------------------------------------------------------
            
        Case TipoDocumento.Envios
            aTexto = "Reparto de Mercadería."
            aComentario = "Código de Impresión Nº: " & Codigo
            '------------------------------------------------------------------------------------------------------------------
        
        Case TipoDocumento.ArregloStock: aTexto = "Arreglo de Stock Automático"
        
        Case TipoDocumento.CambioEstadoMercaderia, TipoDocumento.IngresoMercaderiaEspecial
            Cons = "Select * From ControlMercaderia Where CMeCodigo = " & Codigo
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
            If Not RsAux.EOF Then
                If RsAux!CMeTipo = TipoControlMercaderia.CambioEstado Then
                    aTexto = "Cambio Estado Mercadería"
                Else
                    aTexto = "Ingreso Especial"
                End If
                If Not IsNull(RsAux!CMeComentario) Then aComentario = Trim(RsAux!CMeComentario)
            Else
                If TipoDocumento.CambioEstadoMercaderia = Tipo Then aTexto = "Cambio Estado Mercaderia" Else aTexto = "Ingreso Especial"
            End If
            RsAux.Close
            
            
        Case TipoDocumento.Remito
            Dim aDoc As Long
            Cons = "Select * from Remito Where RemCodigo =  " & Codigo
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            aTexto = "Remito Nº " & Codigo & " / "
            aDoc = RsAux!RemDocumento
            RsAux.Close
            
            Cons = "Select * from Documento, Sucursal" _
                    & " Where DocCodigo =  " & aDoc _
                    & " And DocSucursal = SucCodigo"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            aTexto = aTexto & Trim(RsAux!SucAbreviacion) & " "
            
            Select Case RsAux!DocTipo
                Case TipoDocumento.Contado: aTexto = aTexto & "Contado "
                Case TipoDocumento.Credito: aTexto = aTexto & "Crédito "
            End Select
            
            aTexto = aTexto & Trim(RsAux!DocSerie) & " " & RsAux!DocNumero & " " _
                          & "(" & Format(RsAux!DocFecha, "d/mm/yy hh:mm") & ")"
            RsAux.Close
        
        Case TipoDocumento.Servicio, TipoDocumento.ServicioCambioEstado
            aTexto = "Servicio: " & Codigo
            
        Case TipoDocumento.Devolucion
            aComentario = ""
            aTexto = RetornoNombreDocumento(TipoDocumento.Devolucion) & ": " & Codigo
            Cons = "Select * From Devolucion " _
                        & " Left Outer Join Documento ON DevNota = DocCodigo " _
                & " Where DevID = " & Codigo
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux.EOF Then
                If Not IsNull(RsAux!DevComentario) Then aComentario = Trim(RsAux!DevComentario)
                If Not IsNull(RsAux!DocCodigo) Then
                    If Trim(aComentario) <> "" Then aComentario = aComentario & Chr(13) & Chr(10)
                    aComentario = aComentario & RetornoNombreDocumento(RsAux!DocTipo) & ": " & Trim(RsAux!DocSerie) & " " & Trim(RsAux!DocNumero)
                    If Not IsNull(RsAux!DocComentario) Then aComentario = aComentario & Chr(13) & Chr(10) & "Comentario Nota: " & Trim(RsAux!DocComentario)
                End If
            End If
            RsAux.Close
        
        Case Else
            'Documentos internos de la empresa-------------------------------------------------------------------------
            Cons = "Select * from Documento, Sucursal" _
                    & " Where DocCodigo =  " & Codigo _
                    & " And DocSucursal = SucCodigo"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux.EOF Then
                aTexto = Trim(RsAux!SucAbreviacion) & " "
                
                Select Case RsAux!DocTipo
                    Case TipoDocumento.Contado: aTexto = aTexto & "Contado "
                    Case TipoDocumento.Credito: aTexto = aTexto & "Crédito "
                    Case TipoDocumento.NotaCredito: aTexto = aTexto & "N.Crédito "
                    Case TipoDocumento.NotaDevolucion: aTexto = aTexto & "N.Devolución"
                    Case TipoDocumento.NotaEspecial: aTexto = aTexto & "N.Especial"
                End Select
                
                aTexto = aTexto & Trim(RsAux!DocSerie) & " " & RsAux!DocNumero & " " _
                              & "(" & Format(RsAux!DocFecha, "d/mm/yy hh:mm") & ")"
                
                If Not IsNull(RsAux!DocComentario) Then aComentario = Trim(RsAux!DocComentario)
            '------------------------------------------------------------------------------------------------------------------
                RsAux.Close
                
                If Tipo = TipoDocumento.NotaCredito Or Tipo = TipoDocumento.NotaDevolucion Or Tipo = TipoDocumento.NotaEspecial Then
                    'Levanto los datos del retiro por devolucion
                    Cons = "Select * From Devolucion Where DevNota = " & Codigo _
                        & " And DevFAltaLocal = '" & Format(lFecha.Tag, sqlFormatoFH) & "'" _
                        & " And DevArticulo = " & IdArticulo
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    If Not RsAux.EOF Then
                        If Trim(aComentario) <> "" Then aComentario = aComentario & Chr(13) & Chr(10)
                        aComentario = aComentario & "Devolución: " & RsAux!DevID
                    End If
                    RsAux.Close
                ElseIf (Tipo = TipoDocumento.Contado Or Tipo = TipoDocumento.Credito) And Val(lValor.Caption) >= 1 Then
                    Cons = "Select * From Devolucion " _
                            & " Left Outer Join  Documento ON DevNota = DocCodigo " _
                        & " Where DevFactura = " & Codigo _
                        & " And DevFAltaLocal = '" & Format(lFecha.Tag, sqlFormatoFH) & "'" _
                        & " And DevArticulo = " & IdArticulo
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    If Not RsAux.EOF Then
                        If Trim(aComentario) <> "" Then aComentario = aComentario & Chr(13) & Chr(10)
                        aComentario = aComentario & "Devolución: " & RsAux!DevID
                        'Veo si hizo la nota.
                        If Not IsNull(RsAux!DocCodigo) Then
                            aComentario = aComentario & Chr(13) & Chr(10)
                            aComentario = aComentario & RetornoNombreDocumento(RsAux!DocTipo, True) & " " & RsAux!DocSerie & " " & RsAux!DocNumero
                        End If
                    End If
                    RsAux.Close
                End If
            Else
                RsAux.Close
            End If
            
    End Select
    
    lDocumento.Caption = aTexto
    lTexto.Caption = aComentario
    Exit Sub

errDoc:
    clsGeneral.OcurrioError "Ocurrio un error al cargar los datos del documento.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoDatosArticulo(Codigo As Long)

    On Error GoTo errDoc
    aTexto = ""
    
    Cons = "Select * from Articulo Where ArtId = " & Codigo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    aTexto = Trim(RsAux!ArtNombre)
    RsAux.Close
        
    lArticulo.Caption = aTexto
    Exit Sub

errDoc:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Forms(Forms.Count - 2).SetFocus
End Sub

Private Sub LimpioFicha()
    lFecha.Caption = "S/D"
    lSucursal.Caption = "S/D"
    lUsuario.Caption = "S/D"
    lTerminal.Caption = "S/D"
    lValor.Caption = "S/D"
    lArticulo.Caption = "S/D"
    lDocumento.Caption = "S/D"
    lTexto.Caption = ""
End Sub

Private Function BuscoUsuario(Codigo As Long, Optional Identificacion As Boolean = False, Optional Digito As Boolean = False, Optional Iniciales As Boolean = False)
Dim RsUsr As rdoResultset
Dim aRetorno As String: aRetorno = ""
    
    On Error Resume Next
    
    Cons = "Select * from Usuario Where UsuCodigo = " & Codigo
    Set RsUsr = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsUsr.EOF Then
        If Identificacion Then aRetorno = Trim(RsUsr!UsuIdentificacion)
        If Digito Then aRetorno = Trim(RsUsr!UsuDigito)
        If Iniciales Then aRetorno = Trim(RsUsr!UsuInicial)
    End If
    RsUsr.Close
    
    BuscoUsuario = aRetorno
    
End Function


