VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Begin VB.Form frmAnulacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anulación de Documentos"
   ClientHeight    =   4800
   ClientLeft      =   2445
   ClientTop       =   2610
   ClientWidth     =   7920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAnulacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7920
   Begin AACombo99.AACombo cMotivos 
      Height          =   315
      Left            =   840
      TabIndex        =   10
      Top             =   4020
      Width           =   5535
      _ExtentX        =   9763
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
   Begin AACombo99.AACombo cSucursal 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      BackColor       =   16777215
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
      Text            =   ""
   End
   Begin VB.TextBox tUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   7200
      MaxLength       =   3
      TabIndex        =   12
      Top             =   4020
      Width           =   615
   End
   Begin VB.CommandButton bCancelar 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   5820
      TabIndex        =   13
      Top             =   4440
      Width           =   900
   End
   Begin VB.CommandButton bAnular 
      Caption         =   "&Grabar"
      Height          =   315
      Left            =   6900
      TabIndex        =   14
      Top             =   4440
      Width           =   900
   End
   Begin ComctlLib.ListView lArticulo 
      Height          =   1845
      Left            =   120
      TabIndex        =   7
      Top             =   2070
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   3254
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   327682
      SmallIcons      =   "ImageList1"
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Cantidad"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Artículo"
         Object.Width           =   8114
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Importe"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.TextBox tSerie 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4920
      MaxLength       =   1
      TabIndex        =   5
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox tNumero 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5400
      MaxLength       =   7
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin ComctlLib.ListView lRecibo 
      Height          =   1845
      Left            =   120
      TabIndex        =   8
      Top             =   2070
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   3254
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Factura"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Cuota"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Amortización"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Mora"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Total"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Modificado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "De Cuota"
         Object.Width           =   0
      EndProperty
   End
   Begin AACombo99.AACombo cDocumento 
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      BackColor       =   16777215
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
      Text            =   ""
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Motivos:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Importe:"
      Height          =   255
      Left            =   1920
      TabIndex        =   29
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lImporte 
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2640
      TabIndex        =   28
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario:"
      Height          =   255
      Left            =   6540
      TabIndex        =   11
      Top             =   4065
      Width           =   615
   End
   Begin VB.Label lTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ARTÍCULOS"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   120
      TabIndex        =   27
      Top             =   1800
      Width           =   7680
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Titular:"
      Height          =   255
      Left            =   2280
      TabIndex        =   26
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lTitular 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Rodriguez Fernandez, Rodrigo Bernardino"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2880
      TabIndex        =   25
      Top             =   1440
      UseMnemonic     =   0   'False
      Width           =   4815
   End
   Begin VB.Label lCiRuc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "21 378350 0011"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   960
      TabIndex        =   24
      Top             =   1440
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CI/RUC:"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lComentario 
      BackStyle       =   0  'Transparent
      Caption         =   "Crédito"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1320
      TabIndex        =   22
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Label lMoneda 
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   960
      TabIndex        =   21
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Carlos"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5400
      TabIndex        =   20
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "14-Ene-1998"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6120
      TabIndex        =   19
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Moneda:"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Comentarios:"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Sucursal"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Emisión"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6120
      TabIndex        =   15
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Número"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Tipo de Documento"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   1575
      Left            =   120
      Top             =   120
      Width           =   7695
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuCancelar 
         Caption         =   "&Cancelar"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuGrabar 
         Caption         =   "&Grabar"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu MnuDetalle 
      Caption         =   "&Detalle"
      Begin VB.Menu MnuFactura 
         Caption         =   "&Factura"
         Shortcut        =   ^F
      End
      Begin VB.Menu MnuOperacion 
         Caption         =   "&Operación"
         Shortcut        =   ^O
      End
      Begin VB.Menu MnuPagos 
         Caption         =   "&Pagos"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu MnuVolver 
      Caption         =   "&Salir"
      Begin VB.Menu MnuSalir 
         Caption         =   "Del formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmAnulacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCliente As clsClienteCFE
'Dim oCnfgCtdo As New clsImpresoraTicketsCnfg
Dim oCnfgRecibo As New clsImpresoraTicketsCnfg

Public prmIDDocumento As Long
Private EmpresaEmisora As clsClienteCFE
Dim ProdInteresMora As clsProducto

Dim gDocumento As Long              'Codigo del Documento seleccionado
Dim gCliente As Long                    'Codigo del cliente del Documento
Dim gFechaDocumento As String   'Fecha de Modificacion del Documento Seleccionado
Dim gFModificacionDocQAnulo As String
Dim gFactura As Long                  'Codigo de la Factura/Contado al que se asocia el documento
Dim gTipoPago As Integer            'Tipo de Pago de la Factura (se usa en anulaciones de Recibos)

Dim prmPuedoAnular As Boolean, prmDocAnulado As Boolean
Dim aMensaje As String
Dim aFletes As String           'Cargo los articulos de fletes

Dim jobnum As Integer, CantForm As Integer
Dim bDocDeServicio As Boolean, bQARetirarDif As Boolean

Dim bNotaFleteEnvio As Boolean

Dim itmX As ListItem
Dim aTexto As String, txtERROR As String

Private Sub AnularNotaDeDebito()
    
End Sub


Private Sub bAnular_Click()

    If Trim(cMotivos.Text) = "" Then
        MsgBox "Debe ingresar el motivo de la anulación.", vbExclamation, "Falta Ingreso del Motivos"
        cMotivos.SetFocus: Exit Sub
    End If
    
    If Val(tUsuario.Tag) = 0 Then
        MsgBox "Ingrese el dígito de usuario para anular el documento.", vbExclamation, "ATENCIÓN"
        Foco tUsuario: Exit Sub
    End If
    
    If Not ValidoCajaCerrada Then Exit Sub
    
    If (lArticulo.ListItems.Count = 0 And lRecibo.ListItems.Count = 0) Then
        
        'Veo si lo que se quiere anular es una seña ----------------------------------------------------------------------------------------------------
        If gDocumento <> 0 Then
            Cons = "Select * from Documento, CuentaDocumento " _
                    & " Where DocCodigo = " & gDocumento _
                    & " And DocTipo = " & TipoDocumento.ReciboDePago _
                    & " And DocAnulado = 0" _
                    & " And DocCodigo = CDoIDDocumento"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux.EOF Then
                RsAux.Close
                If MsgBox("Lo que ud. desea anular es un recibo de seña ?." & Chr(vbKeyReturn) _
                            & "Verifique que el saldo de la cuenta, a la que está asignado el recibo, sea mayor o igual al importe del recibo." & Chr(vbKeyReturn) _
                            & "Para anular el recibo de seña presione 'Si'.", vbQuestion + vbYesNo + vbDefaultButton2, "Recibo de Seña") = vbNo Then Exit Sub
                AnuloReciboDeSenia
                Exit Sub
            Else
                RsAux.Close
            End If
        End If
        '------------------------------------------------------------------------------------------------------------------------------------------------------
        
        MsgBox "Debe seleccionar el documento a anular para cargar los datos." & Chr(vbKeyReturn) & "Verifique si este documento no ha sido anulado.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    If MsgBox("Confirma anular el documento seleccionado.", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
    
    Select Case cDocumento.ItemData(cDocumento.ListIndex)
        Case TipoDocumento.Contado: AnuloContado
        Case TipoDocumento.Credito: AnuloCredito
        Case TipoDocumento.ReciboDePago: AnuloReciboDePago
        Case TipoDocumento.NotaDevolucion: AnuloNotaDevolucion
        
        Case TipoDocumento.NotaCredito:
            If Not bNotaFleteEnvio Then AnuloNotaCredito Else AnuloNotaCreditoEnvio
            
        Case TipoDocumento.NotaEspecial: AnuloNotaEspecial
        
        Case TipoDocumento.Remito:
            AnuloRemito
        
    End Select
    
    If prmIDDocumento > 0 Then Unload Me
    
End Sub

Private Sub bCancelar_Click()

    LimpioFicha
    Foco cSucursal
    
End Sub

Private Sub cDocumento_GotFocus()
    cDocumento.SelStart = 0
    cDocumento.SelLength = Len(cDocumento.Text)
End Sub

Private Sub cDocumento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tSerie
End Sub

Private Sub cDocumento_LostFocus()

    cDocumento.SelLength = 0
    
    If cDocumento.ListIndex <> -1 Then
        If cDocumento.ItemData(cDocumento.ListIndex) = TipoDocumento.ReciboDePago Then
            If lTitulo.Tag <> "R" Then
                lTitulo.Caption = " FACTURAS PAGAS      [F] Detalle de Factura    [D] Detalle de la Operación    [P] Ver Pagos"
                lTitulo.Tag = "R"
                lRecibo.ZOrder 0
            End If
        Else
            If lTitulo.Tag <> "A" Then
                lTitulo.Caption = " ARTÍCULOS"
                lTitulo.Tag = "A"
                lArticulo.ZOrder 0
            End If
        End If
    End If
    
End Sub

Private Sub cMotivos_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then If tUsuario.Enabled Then Foco tUsuario
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
        
    FechaDelServidor
    LimpioFicha
    
    'oCnfgCtdo.CargarConfiguracion "FacturaContado", "TicketContado"
    oCnfgRecibo.CargarConfiguracion "ImpresionDocumentos", "ReciboPagoCaja"
    
    If oCnfgRecibo.Opcion = 0 Then
        MsgBox "Debe configurar donde imprimir los eRecibos en el programa de cobranza de cuotas.", vbExclamation, "Atención"
        End
    End If
    
    Set EmpresaEmisora = New clsClienteCFE
    EmpresaEmisora.CargoInformacionCliente cBase, 1, False
    
    CargoArticuloInteresesPorMora
    
    Cons = "Select CAnID, CAnTexto from ComentarioAnulacion Order by CAnTexto"
    CargoCombo Cons, cMotivos
    
    'CentroForm Me
    'Cargo Sucursales---------------------------------------------------------------------------
    Cons = "Select SucCodigo, SucAbreviacion from Sucursal Where SucDisponibilidad is not null Order by SucAbreviacion"
    CargoCombo Cons, cSucursal, ""
    BuscoCodigoEnCombo cSucursal, paCodigoDeSucursal
    
    'Cargo Documentos--------------------------------------------------------------------------
    With cDocumento
'        .AddItem Trim(DocContado): .ItemData(.NewIndex) = TipoDocumento.Contado
        '.AddItem Trim(DocCredito): .ItemData(.NewIndex) = TipoDocumento.Credito
'        .AddItem Trim(DocNCredito): .ItemData(.NewIndex) = TipoDocumento.NotaCredito
'        .AddItem Trim(DocNDevolucion): .ItemData(.NewIndex) = TipoDocumento.NotaDevolucion
'        .AddItem Trim(DocNEspecial): .ItemData(.NewIndex) = TipoDocumento.NotaEspecial
        .AddItem Trim(DocRecibo): .ItemData(.NewIndex) = TipoDocumento.ReciboDePago
'        .AddItem Trim(DocRemito): .ItemData(.NewIndex) = TipoDocumento.Remito
    End With
    '-----------------------------------------------------------------------------------------------
    
    aFletes = CargoArticulosDeFlete
    
    crAbroEngine
    
    If prmIDDocumento > 0 Then
        Me.Show: Me.Refresh
        ProcesoActivacion
    End If
    
End Sub

Private Sub ProcesoActivacion()
On Error GoTo errActivo
    
    Screen.MousePointer = 11
    Cons = "Select * from Documento Where DocCodigo = " & prmIDDocumento
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        BuscoCodigoEnCombo cSucursal, RsAux!DocSucursal
        BuscoCodigoEnCombo cDocumento, RsAux!DocTipo
        tSerie.Text = Trim(RsAux!DocSerie)
        tNumero.Text = RsAux!DocNumero
    End If
    RsAux.Close
    
    If Trim(tNumero.Text) <> "" Then
        Call cDocumento_LostFocus
        Call tNumero_KeyPress(vbKeyReturn)
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errActivo:
    clsGeneral.OcurrioError "Error al cargar los datos del documento.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoArticulos(Documento As Long)
    
    On Error GoTo errCargar
    lArticulo.ListItems.Clear
    bQARetirarDif = False
    
    Screen.MousePointer = 11
    
    Cons = "Select ArtID, ArtNombre, RenCantidad, RenARetirar, RenPrecio" _
           & " From Renglon, Articulo" _
           & " Where RenDocumento = " & Documento _
           & " And RenArticulo = ArtID"

    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        
        Set itmX = lArticulo.ListItems.Add(, "A" & Str(RsAux!ArtID), RsAux!RenCantidad)
        itmX.SubItems(1) = Trim(RsAux!ArtNombre)
        itmX.SubItems(2) = Format(RsAux!RenPrecio * RsAux!RenCantidad, FormatoMonedaP)
        
        If RsAux!RenCantidad <> RsAux!RenARetirar Then
            If InStr(aFletes, "," & RsAux!ArtID & ",") = 0 Then
                prmPuedoAnular = False
                bQARetirarDif = True
            End If
        End If
        RsAux.MoveNext
    Loop
    
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    prmPuedoAnular = False
    clsGeneral.OcurrioError "Error al cargar los artículos del documento.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Screen.MousePointer = 11
    crCierroEngine
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    
    Screen.MousePointer = 0
    End
    
End Sub

Private Sub Label1_Click()
    Foco cDocumento
End Sub

Private Sub Label13_Click()
    Foco tUsuario
End Sub

Private Sub Label9_Click()
    Foco cSucursal
End Sub

Private Sub cSucursal_GotFocus()
    cSucursal.SelStart = 0
    cSucursal.SelLength = Len(cSucursal.Text)
End Sub

Private Sub cSucursal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cDocumento
End Sub

Private Sub cSucursal_LostFocus()
    cSucursal.SelLength = 0
End Sub

Private Sub CargoCliente(Cliente As Long)
    
    On Error GoTo errCliente
'     Cons = "Select CliCiRuc, CliTipo, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2) " _
'           & " From Cliente, CPersona " _
'           & " Where CliCodigo = " & Cliente _
'           & " And CliCodigo = CPeCliente " _
'           & " UNION " _
'           & " Select CliCiRuc, CliTipo, Nombre = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')')  " _
'           & " From Cliente, CEmpresa " _
'           & " Where CliCodigo = " & Cliente _
'           & " And CliCodigo = CEmCliente"
'
'    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    lCiRuc.Caption = "S/D"
'    If Not RsAux.EOF Then
'        lCiRuc.Tag = RsAux!CliTipo
'        If RsAux!CliTipo = 1 Then
'            If Not IsNull(RsAux!CliCIRuc) Then lCiRuc.Caption = clsGeneral.RetornoFormatoCedula(RsAux!CliCIRuc)
'        Else
'            If Not IsNull(RsAux!CliCIRuc) Then lCiRuc.Caption = clsGeneral.RetornoFormatoRuc(Trim(RsAux!CliCIRuc))
'        End If
'    End If
'    lTitular.Caption = Trim(RsAux!Nombre)
'    RsAux.Close

    Set oCliente = New clsClienteCFE
    oCliente.CargoInformacionCliente cBase, Cliente, False
    lCiRuc.Tag = oCliente.TipoCliente
    lTitular.Caption = oCliente.NombreCliente
    If oCliente.TipoCliente = TC_Persona Then
        If oCliente.RUT <> "" Then
            lCiRuc.Caption = clsGeneral.RetornoFormatoRuc(oCliente.RUT): lTitular.Tag = "1"
        Else
            lCiRuc.Caption = clsGeneral.RetornoFormatoCedula(oCliente.CI): lTitular.Tag = ""
        End If
    Else
        If oCliente.RUT <> "" Then lCiRuc.Caption = clsGeneral.RetornoFormatoRuc(oCliente.RUT): lTitular.Tag = "1"
    End If
    Exit Sub
    
errCliente:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub lArticulo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Foco tUsuario
    
End Sub

Private Sub lRecibo_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyD     'Detalle de la operacion
            If lRecibo.ListItems.Count = 0 Then Exit Sub
            EjecutarApp prmPathApp & "\Detalle de operaciones", CLng(lRecibo.SelectedItem.Tag)
        
        Case vbKeyP     'Ver Pagos
            If lRecibo.ListItems.Count = 0 Then Exit Sub
            EjecutarApp prmPathApp & "\Detalle de pagos", CLng(lRecibo.SelectedItem.Tag)
            
        Case vbKeyF
            If lRecibo.ListItems.Count = 0 Then Exit Sub
            EjecutarApp prmPathApp & "\Detalle de factura", CLng(lRecibo.SelectedItem.Tag)
            
    End Select
    
End Sub

Private Sub lRecibo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tUsuario
End Sub

Private Sub MnuCancelar_Click()
    Call bCancelar_Click
End Sub

Private Sub MnuFactura_Click()
    Dim aIdDocumento As Long
    
    If lRecibo.ListItems.Count = 0 Then aIdDocumento = gFactura Else aIdDocumento = CLng(lRecibo.ListItems(1).Tag)
    EjecutarApp prmPathApp & "\Detalle de factura", CStr(aIdDocumento)
    
End Sub

Private Sub MnuGrabar_Click()
    Call bAnular_Click
End Sub

Private Sub MnuOperacion_Click()
Dim aIdDocumento As Long
    
    If lRecibo.ListItems.Count = 0 Then aIdDocumento = gFactura Else aIdDocumento = CLng(lRecibo.ListItems(1).Tag)
    
    EjecutarApp prmPathApp & "\Detalle de operaciones", CStr(aIdDocumento)
End Sub

Private Sub MnuPagos_Click()
    EjecutarApp prmPathApp & "\Detalle de pagos", CStr(gFactura)
End Sub

Private Sub MnuSalir_Click()
    Unload Me
End Sub

Private Sub tNumero_GotFocus()
    tNumero.SelStart = 0
    tNumero.SelLength = Len(tNumero.Text)
End Sub

Private Sub tNumero_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        bDocDeServicio = False
        
        If cDocumento.ListIndex = -1 Or cSucursal.ListIndex = -1 Or Trim(tSerie.Text) = "" _
            Or Not IsNumeric(tNumero.Text) Or IsNumeric(tSerie.Text) Then
            MsgBox "Los datos ingresados no son correctos. Verifique.", vbExclamation, "ATENCIÓN"
             Exit Sub
        End If
        
        On Error GoTo errCargar
        Screen.MousePointer = 11
        LimpioFicha
        
        Select Case cDocumento.ItemData(cDocumento.ListIndex)
            Case TipoDocumento.Contado: CargoDatosContado
            Case TipoDocumento.Credito: CargoDatosCredito
            Case TipoDocumento.ReciboDePago: CargoDatosReciboPago
            Case TipoDocumento.NotaDevolucion: CargoDatosNotaDevolucion
            Case TipoDocumento.NotaCredito: CargoDatosNotaCredito
            
            Case TipoDocumento.NotaEspecial: CargoDatosNotaEspecial
            Case TipoDocumento.Remito: CargoDatosRemito
        End Select
'        If tUsuario.Enabled Then Foco tUsuario
        If cMotivos.Enabled Then cMotivos.SetFocus
        Screen.MousePointer = 0
    End If
    Exit Sub

errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del documento."
End Sub

Private Function fnc_CargoDocumento(Tipo As Integer) As Boolean
'   Si esto retorna FALSO quiere decir que ya no lo puedo ANULAR !!!

    On Error GoTo errCargar
    fnc_CargoDocumento = True
    
    gFactura = 0
    prmPuedoAnular = True
    prmDocAnulado = False
    
    Cons = "Select * From Documento Left Outer Join Usuario on DocUsuario = UsuCodigo " _
                                                 & " Left Outer Join Moneda on DocMoneda = MonCodigo " _
           & " Where DocTipo = " & Tipo _
           & " And DocSerie = '" & Trim(tSerie.Text) & "'" _
           & " And DocNumero = " & Trim(tNumero.Text) _
           & " And DocSucursal = " & cSucursal.ItemData(cSucursal.ListIndex) & " Order by doccodigo desc"
           
           
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    gDocumento = 0
    If Not RsAux.EOF Then
        gDocumento = RsAux!doccodigo
        gCliente = RsAux!DocCliente
        gFechaDocumento = RsAux!DocFecha
        
        gFModificacionDocQAnulo = RsAux!DocFModificacion
        
        lFecha.Caption = Format(RsAux!DocFecha, "d-Mmm-yy hh:mm:ss")
        lFecha.Tag = RsAux!DocFecha
        If Not IsNull(RsAux!UsuIdentificacion) Then lUsuario.Caption = Trim(RsAux!UsuIdentificacion)
    
        lComentario.Caption = "N/D"
        If Not IsNull(RsAux!DocComentario) Then lComentario.Caption = Trim(RsAux!DocComentario)
    
        lMoneda.Caption = Trim(RsAux!MonSigno)
        lMoneda.Tag = RsAux!DocMoneda
        lImporte.Caption = Format(RsAux!DocTotal, FormatoMonedaP)
        If Not IsNull(RsAux!Dociva) Then lImporte.Tag = Format(RsAux!Dociva, FormatoMonedaP) Else lImporte.Tag = 0
        
        If RsAux!DocAnulado Then
            MsgBox "El documento seleccionado ya ha sido anulado.", vbInformation, "ATENCIÓN"
            prmPuedoAnular = False
            prmDocAnulado = True
        End If
        
        If RsAux!DocTipo = TipoDocumento.ReciboDePago And RsAux!DocTotal < 0 Then
            If Format(RsAux!DocFecha, "dd/mm/yyyy") <> Format(gFechaServidor, "dd/mm/yyyy") Then
                MsgBox "El recibo de pago no se puede anular." & vbCrLf & _
                            "No es del día y es una devolución sobre otro recibo de pago.", vbInformation, "ATENCIÓN"
                prmPuedoAnular = False
            End If
        End If
        
    Else
        MsgBox "No existe un documento para las características ingresadas.", vbInformation, "ATENCIÓN"
        prmPuedoAnular = False
    End If
    RsAux.Close
    
    fnc_CargoDocumento = prmPuedoAnular
    
    Exit Function
errCargar:
    clsGeneral.OcurrioError "Error al cargar los datos del documento.", Err.Description
    HabilitoCampos False
End Function

Private Sub CargoDatosContado()

    If Not fnc_CargoDocumento(TipoDocumento.Contado) Then Exit Sub
    
    gFactura = gDocumento
    
    If gDocumento = 0 Then Exit Sub
    
    CargoCliente gCliente
    CargoArticulos gDocumento
        
    If prmDocAnulado Then HabilitoCampos False: Exit Sub
        
    If Format(gFechaDocumento, "yyyy/mm/dd") <> Format(gFechaServidor, "yyyy/mm/dd") Then
        MsgBox "El documento no se puede anular." & vbCrLf & _
                    "Sólo se podrán anular los documentos del día.", vbInformation, "El documento NO es del DÍA"
        HabilitoCampos False: Exit Sub
    End If
        
    If HayRemitos Then GoTo Salir
    If HayNota Then GoTo Salir
    If HayEnvio Then GoTo Salir
    HayInstalacion
    HayEnvioCobroFlete True
    
    '(para el control MultiUsuario)
    Cons = "Select * from Documento Where DocCodigo = " & gFactura
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    gFechaDocumento = RsAux!DocFModificacion
    RsAux.Close
            
     If Not prmPuedoAnular And bQARetirarDif Then
        '1) Veo si el doc es de servicio
        DocumentoDeServicios gDocumento
        If bDocDeServicio Then prmPuedoAnular = True
        If Not prmPuedoAnular Then
            aMensaje = "El documento no se puede anular, las cantidades de venta no coinciden con las a retirar."
            GoTo Salir
        End If
    End If
    
'    If prmPuedoAnular Then        'Cambios el 27/03/2003
        'Si el Doc. es de servicio y no cumplio ninguna de las condiciones anterior dejo anular
'        DocumentoDeServicios gDocumento
'    Else
'        aMensaje = "El documento no se puede anular, las cantidades de venta no coinciden con las a retirar."
'        GoTo Salir
'    End If

    Dim rsDP As rdoResultset
    Cons = "SELECT DPeDocumento FROM DocumentoPendiente WHERE DPeDocumento = " & gDocumento
    Set rsDP = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsDP.EOF Then
        MsgBox "ATENCIÓN!!!" & vbCrLf & "El documento está en pendiente de camioneros, verifique que realmente no se va a cobrar el flete o el servicio asociado.", vbExclamation, "PENDIENTE EN CAMIONEROS"
    End If
    rsDP.Close

    If prmPuedoAnular Then HabilitoCampos True
        
    Exit Sub
    
Salir:
    MsgBox aMensaje, vbInformation, "ATENCIÓN"
    HabilitoCampos False
End Sub

Private Sub CargoDatosRemito()

    If Not fnc_CargoDocumento(TipoDocumento.Remito) Then Exit Sub
    
    gFactura = gDocumento
    
    If gDocumento = 0 Then Exit Sub
    
    CargoCliente gCliente
    CargoArticulos gDocumento
        
    If prmDocAnulado Then HabilitoCampos False: Exit Sub
        
    If Format(gFechaDocumento, "yyyy/mm/dd") < Format(DateAdd("d", paDiasAnulacionRemito * -1, gFechaServidor), "yyyy/mm/dd") Then
        MsgBox "El documento no se puede anular." & vbCrLf & _
                    "La fecha del documento es menor a la fecha permitida.", vbInformation, "El documento NO es del DÍA"
        HabilitoCampos False: Exit Sub
    End If
        
    '(para el control MultiUsuario)
    Cons = "Select * from Documento Where DocCodigo = " & gFactura
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    gFechaDocumento = RsAux!DocFModificacion
    RsAux.Close
        
     If Not prmPuedoAnular And bQARetirarDif Then
        '1) Veo si el doc es de servicio
'        DocumentoDeServicios gDocumento
'        If bDocDeServicio Then prmPuedoAnular = True
        If Not prmPuedoAnular Then
            aMensaje = "El documento no se puede anular, las cantidades de venta no coinciden con las a retirar."
            GoTo Salir
        End If
    End If
       
    If prmPuedoAnular Then HabilitoCampos True
    Exit Sub
    
Salir:
    MsgBox aMensaje, vbInformation, "ATENCIÓN"
    HabilitoCampos False
End Sub

Private Sub DocumentoDeServicios(IdDocumento As Long)
Dim rsSer As rdoResultset
    
    On Error GoTo errSer
    'Veo si el documento Factura servicios----------------------------------------------
    Cons = "Select * from Servicio Where SerDocumento = " & IdDocumento
    Set rsSer = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsSer.EOF Then bDocDeServicio = True Else bDocDeServicio = False
    rsSer.Close
    
    If bDocDeServicio Then MsgBox "Este documento factura una reparación en taller o un servicio." & Chr(vbKeyReturn) & "Si ud. lo anula no se producirán movimientos de stock.", vbInformation, "Documento de Servicios."
    
    Exit Sub

errSer:
    clsGeneral.OcurrioError "Error al validar si el documento es de servicios.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoDatosCredito()

    If Not fnc_CargoDocumento(TipoDocumento.Credito) Then Exit Sub
    gFactura = gDocumento
    
    If gDocumento = 0 Then Exit Sub
    
    CargoCliente gCliente
    CargoArticulos gDocumento
    
    If prmDocAnulado Then HabilitoCampos False: Exit Sub
    
    If Format(gFechaDocumento, "yyyy/mm/dd") <> Format(gFechaServidor, "yyyy/mm/dd") Then
        MsgBox "El documento no se puede anular. Sólo se podrán anular los documentos del día.", vbInformation, "ATENCIÓN"
        HabilitoCampos False: Exit Sub
    End If
    
    '(para el control MultiUsuario)
    Cons = "Select * from Documento Where DocCodigo = " & gFactura
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    gFechaDocumento = RsAux!DocFModificacion
    RsAux.Close
    
    'If prmPuedoAnular Then    '27/03/03 Cambio - Veo la causa de no poder anular------------------------------------------------------
    If HayRemitos Then GoTo Salir
    If HayNota Then GoTo Salir
    If HayEnvio Then GoTo Salir
    HayInstalacion
    
    If Not prmPuedoAnular And bQARetirarDif Then
        '1) Veo si el doc es de servicio
        DocumentoDeServicios gDocumento
        If bDocDeServicio Then prmPuedoAnular = True
        If Not prmPuedoAnular Then
            aMensaje = "El documento no se puede anular, las cantidades de venta no coinciden con las a retirar."
            GoTo Salir
        End If
    End If
    
    'Valido si tiene algun Recibo Pago------------------------------------------------------------------------------------------
'    Cons = "Select * from DocumentoPago, Documento" _
'            & " Where DPaDocASaldar = " & gDocumento _
'            & " And DPaDocQSalda = DocCodigo" _
'            & " And DocAnulado = 0"     '0 = Falso
'    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
'    If Not RsAux.EOF Then
'        aMensaje = "El documento no se puede anular debido a que tiene asociado RECIBOS de PAGO"
'        RsAux.Close
'        GoTo Salir
'    End If
'    RsAux.Close
    If CreditoConRecibos() Then GoTo Salir
    '-------------------------------------------------------------------------------------------------------------------------------
    If prmPuedoAnular Then HabilitoCampos True
    
    Exit Sub
Salir:
    MsgBox aMensaje, vbInformation, "ATENCIÓN"
    HabilitoCampos False
End Sub

Private Function CreditoConRecibos() As Boolean
Dim rsCR As rdoResultset
    CreditoConRecibos = False
    Cons = "SELECT DocCodigo from DocumentoPago, Documento" _
        & " WHERE DPaDocASaldar = " & gDocumento _
        & " AND DPaDocQSalda = DocCodigo AND DocAnulado = 0"
    Set rsCR = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsCR.EOF Then
        aMensaje = "El documento no se puede anular debido a que tiene asociado RECIBOS de PAGO"
        CreditoConRecibos = True
    End If
    rsCR.Close
End Function

Private Function TengoRecibosPosteriores(ByVal idFactura As Long, ByVal idRecibo As Long) As Boolean

    Cons = "Select IsNull(Sum(DocTotal), 0) from DocumentoPago, Documento" _
            & " Where DPaDocASaldar = " & idFactura _
            & " And DPaDocQSalda <> " & idRecibo _
            & " And DPaDocQSalda = DocCodigo" _
            & " And DocCodigo > " & gDocumento _
            & " And DocAnulado = 0"     '0 = Falso
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    TengoRecibosPosteriores = (RsAux(0) <> 0)
    RsAux.Close

End Function

Private Sub CargoDatosReciboPago()

    If Not fnc_CargoDocumento(TipoDocumento.ReciboDePago) Then Exit Sub
    Dim txtNotaDebito As String
    
    If gDocumento <> 0 Then
    
        '29/05/2012 creo var y voy sumando mora, luego como la nota de débito puede ser la del final del día aplico sólo el porcentaje de este recibo.
        Dim sumoMora As Currency
        Dim ivaPorcNota As Currency
        
        CargoCliente gCliente
                
        'Cargo las facturas a las que salda------------------------------------------------------------------------------------------
        Cons = "Select * from DocumentoPago, Documento" _
                & " Where DPaDocQSalda = " & gDocumento _
                & " And DPaDocASaldar = DocCodigo" _
                & " And DocAnulado = 0 " _
                & " Order by DPaCuota DESC" _
                '& " Order by DPaDocASaldar, DPaDocQSalda, DPaCuota " 'DESC"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not RsAux.EOF
            
                Set itmX = lRecibo.ListItems.Add(, , Trim(RsAux!DocSerie) & Trim(RsAux!DocNumero))
                itmX.Key = Chr((Asc("A") + lRecibo.ListItems.Count - 1)) & RsAux!DocTipo
                
                itmX.Tag = RsAux!doccodigo  'Codigo de la Factura
                
                If RsAux!DPaCuota = 0 Then itmX.SubItems(1) = "E" Else: itmX.SubItems(1) = Trim(RsAux!DPaCuota)
                If Not IsNull(RsAux!DPaAmortizacion) Then itmX.SubItems(2) = Format(RsAux!DPaAmortizacion, FormatoMonedaP) Else itmX.SubItems(2) = "0.00"
                If IsNull(RsAux!DPaMora) Then itmX.SubItems(3) = "0.00" Else: itmX.SubItems(3) = Format(RsAux!DPaMora, FormatoMonedaP): sumoMora = sumoMora + RsAux("DPaMora")
                
                itmX.SubItems(4) = Format(CCur(itmX.SubItems(2)) + CCur(itmX.SubItems(3)), FormatoMonedaP)
                itmX.SubItems(5) = Format(RsAux!DocFModificacion, "dd/mm/yyyy hh:mm:ss")
                itmX.SubItems(6) = RsAux!DPaDe
            
                If (RsAux!DocTipo = TipoDocumento.NotaDebito) Then
                    txtNotaDebito = txtNotaDebito & IIf(txtNotaDebito = "", "", "; ") & Trim(RsAux!DocSerie) & "-" & RsAux!DocNumero
'                    'Como es un Recibo, si se cobró mora el IVA está en las NOTAS por lo tanto lo calculo acá (xsi al anular hay que hacer recibo en negativo)
'                    If Not IsNumeric(lImporte.Tag) Then lImporte.Tag = 0
'                    If Not IsNull(RsAux!DocIVA) Then lImporte.Tag = Format(CCur(lImporte.Tag) + RsAux!DocIVA, FormatoMonedaP)
'                        ' como sacar el porcentaje Format((RsAux("DocIVA") * 100) / (RsAux("DocTotal") - RsAux("DocIVA")), "##")
'                    End If
                End If
            
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        
        '29/05/2012 el recibo NO LLEVA IVA sólo la nota de débito.
        lImporte.Tag = 0
        '-------------------------------------------------------------------------------------------------------------------------------
        
        If prmDocAnulado Then HabilitoCampos False: Exit Sub
        
        'Si la factura fue paga c/cheques no se puede anular------------------------
        If lRecibo.ListItems.Count > 0 Then
            If FacturaPagaConCheques(CLng(lRecibo.ListItems(1).Tag), Recibo:=True) Then
                MsgBox aMensaje, vbExclamation, "ATENCIÓN"
                HabilitoCampos False
                Exit Sub
            End If
        End If
        
        'Verifico si hay Nota para algunos de los documentos---------------------------------------------------------------
        For Each itmX In lRecibo.ListItems
            Cons = "Select * from Nota, Documento " _
                    & " Where NotFactura = " & CLng(itmX.Tag) _
                    & " And NotNota = DocCodigo" _
                    & " And DocAnulado = 0" _
                    & " And DocTipo =" & TipoDocumento.NotaCredito  'Agregue esto el 2007/01/09
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux.EOF Then
                If RsAux!DocFecha > CDate(gFechaDocumento) Then
                    aMensaje = "El documento no se puede anular debido a que la FACTURA " & Trim(itmX.Text) & " tiene asociada la NOTA Nº " & Trim(RsAux!DocSerie) & Trim(RsAux!DocNumero)
                    RsAux.Close
                    GoTo Salir
                    Exit For
                End If
            End If
            RsAux.Close
        Next
        '-----------------------------------------------------------------------------------------------------------------------------
        Dim bNotaMultiple As Boolean
        If ValidoReciboEnNotaMultiple(gDocumento) Then
            bNotaMultiple = True
            MsgBox "Atención!!!" & vbCrLf & vbCrLf & "El recibo está en una nota de débito múltiple posterior a su emisión, verifique.", vbExclamation, "ATENCIÓN"
            'Exit Sub
'            aMensaje = "Ud. no puede anular el Recibo de Pago porque está asignado a una Nota de Débito múltiple."
'            GoTo Salir
        End If
        '-----------------------------------------------------------------------------------------------------------------------------
        
        
        For Each itmX In lRecibo.ListItems
            If CLng(Mid(itmX.Key, 2, Len(itmX.Key))) <> 40 Then ' And Not bNotaMultiple Then
                If TengoRecibosPosteriores(CLng(itmX.Tag), gDocumento) Then
                    aMensaje = "El documento no se puede anular debido a que la FACTURA " & Trim(itmX.Text) & " tiene asociados RECIBOS de pagos posteriores."
                    GoTo Salir
                    Exit Sub
                End If
            End If
        Next
        
        '-----------------------------------------------------------------------------------------------------------------------------
        
        
        If prmPuedoAnular Then
            HabilitoCampos True
            If txtNotaDebito <> "" And Not bNotaMultiple Then
                MsgBox "El recibo ingresado incluye cobranza de moras en la Nota de Débito " & txtNotaDebito & vbCrLf & _
                            "Si usted anula el recibo, se emitirá una nota de débito negativa para anular la misma.", vbInformation, "Anular Recibos"
            End If
        End If
    End If
    Exit Sub
    
Salir:
    MsgBox aMensaje, vbInformation, "ATENCIÓN"
    HabilitoCampos False
    Exit Sub
End Sub

Private Sub CargoDatosNotaDevolucion()

    If Not fnc_CargoDocumento(TipoDocumento.NotaDevolucion) Then Exit Sub
    If gDocumento <> 0 Then
        CargoCliente gCliente
        CargoArticulos gDocumento
                
        'Controlo si es del dia con la fecha de la nota
        If Format(gFechaDocumento, "yyyy/mm/dd") <> Format(gFechaServidor, "yyyy/mm/dd") Then
            MsgBox "El documento no se puede anular. Sólo se podrán anular los documentos del día.", vbInformation, "ATENCIÓN"
            HabilitoCampos False: Exit Sub
        End If
        
        'Busco el Código de la Factura asociada a la Nota
        gFactura = CodigoDeFacturaEnNota(gDocumento)
        
        VerificoDevoluciones gDocumento
        
        'Como es una Nota la gFechaDocumento va a ser la de la Factura (para el control MultiUsuario)
        Cons = "Select * from Documento Where DocCodigo = " & gFactura
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        gFechaDocumento = RsAux!DocFModificacion
        RsAux.Close
        '------------------------------------------------------------------------------------------------------------
        
        If prmDocAnulado Then
            HabilitoCampos False
        Else
        '17/10/2003 Saco este control pq no sabiamos c/carlos px lo hicimos
        'Habia un nota de flete y otra hecha dps, Carlos queria anular y no podia, Sacamos el control
        
        '    If HayNotasParaFactura(CDate(Trim(lFecha.Tag)), gFactura) Then
        '        MsgBox aMensaje, vbExclamation, "ATENCIÓN"
        '        HabilitoCampos False
        '    Else
                HabilitoCampos True
        '    End If
        End If
        
    End If
    
End Sub

Private Sub CargoDatosNotaEspecial()

    If Not fnc_CargoDocumento(TipoDocumento.NotaEspecial) Then Exit Sub
    If gDocumento <> 0 Then
        CargoCliente gCliente
        CargoArticulos gDocumento
                
        'Controlo si es del dia con la fecha de la nota
        If Format(gFechaDocumento, "yyyy/mm/dd") <> Format(gFechaServidor, "yyyy/mm/dd") Then
            MsgBox "El documento no se puede anular. " & vbCrLf & "Sólo se podrán anular los documentos del día.", vbInformation, "ATENCIÓN"
            HabilitoCampos False: Exit Sub
        End If
        
        'Busco el Código de la Factura asociada a la Nota
        gFactura = CodigoDeFacturaEnNota(gDocumento)
        
        VerificoDevoluciones gDocumento
        
        'Como es una Nota la gFechaDocumento va a ser la de la Factura (para el control MultiUsuario)
        If gFactura > 0 Then
            Cons = "Select * from Documento Where DocCodigo = " & gFactura
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            gFechaDocumento = RsAux!DocFModificacion
            RsAux.Close
        End If
        '------------------------------------------------------------------------------------------------------------
        
        If prmDocAnulado Then
            HabilitoCampos False
        Else
            If HayNotasParaFactura(CDate(Trim(lFecha.Tag)), gFactura) Then
                MsgBox aMensaje, vbExclamation, "ATENCIÓN"
                HabilitoCampos False
            Else
                HabilitoCampos True
            End If
        End If
        
    End If
    
End Sub

Private Sub CargoDatosNotaCredito()

    bNotaFleteEnvio = False
    If Not fnc_CargoDocumento(TipoDocumento.NotaCredito) Then Exit Sub
    
    If gDocumento <> 0 Then
        CargoCliente gCliente
        CargoArticulos gDocumento
        
        'Controlo si es del dia con la fecha de la nota
        If Format(gFechaDocumento, "yyyy/mm/dd") <> Format(gFechaServidor, "yyyy/mm/dd") Then
            MsgBox "El documento no se puede anular. Sólo se podrán anular los documentos del día.", vbInformation, "ATENCIÓN"
            HabilitoCampos False: Exit Sub
        End If
        
        VerificoDevoluciones gDocumento
        
        'Busco el Código de la Factura asociada a la Nota
        gFactura = CodigoDeFacturaEnNota(gDocumento)
        
        'Como es una Nota la gFechaDocumento va a ser la de la Factura (para el control MultiUsuario)
        Cons = "Select * from Documento Where DocCodigo = " & gFactura
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        gFechaDocumento = RsAux!DocFModificacion
        RsAux.Close
        '------------------------------------------------------------------------------------------------------------
        
        If prmDocAnulado Then
            HabilitoCampos False
        Else
            If FacturaPagaConCheques(gFactura, Nota:=True) Then
                MsgBox aMensaje, vbExclamation, "ATENCIÓN"
                HabilitoCampos False
                Exit Sub
            End If
            
            If HayRecibosParaFactura(CDate(Trim(lFecha.Tag)), gFactura) Then
                MsgBox aMensaje, vbExclamation, "ATENCIÓN"
                HabilitoCampos False
                Exit Sub
            End If
            
            If HayNotasParaFactura(CDate(Trim(lFecha.Tag)), gFactura) Then
                MsgBox aMensaje, vbExclamation, "ATENCIÓN"
                HabilitoCampos False
                Exit Sub
            End If
            
            'Si solo son Arts. de Fletes no se puede anular la nota produce error al reintegrar el valor, quedam las cuotas con Saldo 0.
            'Es porque los 39 $ estan en la primera y yo intento repartirlo en todas las cuotas 13/9
            Dim bOkFletes As Boolean: bOkFletes = False
            For Each itmX In lArticulo.ListItems
                If InStr(aFletes, Trim(Mid(itmX.Key, 2)) & ",") = 0 Then
                    bOkFletes = True: Exit For
                End If
            Next
            If Not bOkFletes Then
                bNotaFleteEnvio = True
            '    HabilitoCampos False
                MsgBox "Esta nota es de un flete asociado a una factura de crédito y es emitida por el Envío." & vbCrLf & _
                            "Si ud. desea volver a cobrar el flete emita una boleta de contado.", vbExclamation, "Nota sobre Flete"
            End If
            HabilitoCampos True
        End If
    End If
    
End Sub

Private Sub VerificoDevoluciones(idNota As Long)

    'Verifico datos en tabla devolucion----------------------------------------------------------------------------------------
        Dim aStr As String: aStr = ""
        Cons = "Select * From Devolucion, Articulo Where DevNota = " & idNota & " And DevArticulo = ArtId"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            aStr = "Mercadería Devuelta por Nota" & Chr(vbKeyReturn)
            Do While Not RsAux.EOF
                aStr = aStr & RsAux!DevCantidad & Format(RsAux!ArtCodigo, " (#,000,000)") & " " & Trim(RsAux!ArtNombre)
                If Not IsNull(RsAux!DevLocal) Then aStr = aStr & " (en local)" Else aStr = aStr & " (no ingresó)"
                aStr = aStr & Chr(vbKeyReturn)
                
                RsAux.MoveNext
            Loop
            
            aStr = aStr & Chr(vbKeyReturn) & "Si la mercadería ingresó al stock, consulte antes de anular !!!."
        End If
        RsAux.Close
        If aStr <> "" Then MsgBox aStr, vbExclamation, "Mercadería Devuelta."
        Me.Refresh
        '-----------------------------------------------------------------------------------------------------------------------------
End Sub

Private Function CodigoDeFacturaEnNota(Nota As Long) As Long
    
    On Error GoTo errBuscar
    Cons = "Select * from Nota Where NotNota = " & Nota
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    CodigoDeFacturaEnNota = RsAux!NotFactura
    RsAux.Close
    Exit Function
    
errBuscar:
End Function
Private Sub tSerie_GotFocus()
    tSerie.SelStart = 0
    tSerie.SelLength = Len(tSerie.Text)
End Sub

Private Sub tSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then Foco tNumero
End Sub

Private Sub LimpioFicha()
    
    Set oCliente = Nothing
    prmPuedoAnular = True
    bNotaFleteEnvio = False
    bQARetirarDif = False
    
    lComentario.Caption = "S/D"
    lMoneda.Caption = "S/D"
    lImporte.Caption = "S/D"
    lFecha.Caption = "S/D"
    lCiRuc.Caption = "S/D"
    lTitular.Caption = "S/D"
    lTitular.Tag = ""
    lUsuario.Caption = "S/D"
    
    lArticulo.ListItems.Clear
    lRecibo.ListItems.Clear
    
    cMotivos.Text = ""
    tUsuario.Text = ""
    tUsuario.Tag = 0
        
    HabilitoCampos False
    
End Sub

Private Function AnuloRemito() As Boolean
On Error GoTo errAnular

    Screen.MousePointer = 11
    
    Dim obj_DOC As clsFunciones
    Set obj_DOC = New clsFunciones
    
    Set obj_DOC.Connect = cBase
    AnuloRemito = (obj_DOC.AnularRemito(gDocumento, True, True) = 0)
    Set obj_DOC = Nothing
    
    Screen.MousePointer = 0
    If AnuloRemito Then
        LimpioFicha
        Foco cSucursal
    Else
        'En la dll tiene que no despliegue los errores.
        MsgBox "No se logró anular el remito, reintente.", vbExclamation, "Atención"
    End If

    Exit Function
    
errAnular:
    clsGeneral.OcurrioError "Error al anular el Remito de Mercadería", Err.Description
End Function

Private Sub AnuloInstalacion()
    Cons = "UPDATE Instalacion SET InsAnulada = GETDATE(), InsFechaModificacion = GetDATE() WHERE InsTipoDocumento = 1 AND InsDocumento = " & gDocumento & " AND InsAnulada IS NULL"
    cBase.Execute Cons
End Sub
Private Sub AnuloContado()
Dim RsCon As rdoResultset

    Screen.MousePointer = 11
    Cons = "Select * from Documento Where DocCodigo = " & gDocumento
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
    If RsAux!DocFModificacion <> gFechaDocumento Then
        MsgBox "El documento ha sido modificado por otro usuario. Vuelva a cargar los datos.", vbExclamation, "ATENCIÓN"
        RsAux.Close: Screen.MousePointer = 0: Exit Sub
    End If
    
    On Error GoTo errorBT
    Screen.MousePointer = 11
    
    'Verifico que el documento no este asignado a alguna cuenta
    If VerificoCuentaDocumento(gDocumento) Then
        MsgBox "El documento seleccionado está asignado a cuentas o colectivos." & Chr(vbKeyReturn) & _
                    "Para anularlo primero debe quitar la asignación.", vbInformation, "Documento Asignado a Cuentas"
        Screen.MousePointer = 0: Exit Sub
    End If
    
    FechaDelServidor        'Saco la fecha del servidor
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
    
    'Actualizo el documento---------------------------------------------------
    RsAux.Edit
    RsAux!DocFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    RsAux!DocAnulado = True
    RsAux.Update
    '------------------------------------------------------------------------------
    
    'Si el documento está en pendiente de depósito entonces lo elimino del mismo.
    Dim rsDP As rdoResultset
    Cons = "SELECT * FROM DocumentoPendiente WHERE DPeDocumento = " & gDocumento
    Set rsDP = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsDP.EOF Then
        rsDP.Edit
        rsDP("DPeFLiquidacion") = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
        rsDP("DPeIDLiquidacion") = -1
        rsDP.Update
    End If
    rsDP.Close
    
    AnuloInstalacion

    Cons = "Select * from Envio " _
        & " Where EnvTipo = " & TipoEnvio.Entrega _
        & " And EnvDocumentoFactura = " & gDocumento & " AND EnvDocumento <> EnvDocumentoFactura AND EnvEstado <> 4"
    Set rsDP = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsDP.EOF Then
        rsDP.Edit
        rsDP("EnvFormaPago") = paFPagoAnulaDocumento
        rsDP("EnvLiquidar") = 0
        rsDP("EnvDocumentoFactura") = Null
        rsDP.Update
    End If
    rsDP.Close

    'Agrego la Mercadería al Stock-----------------------------------------------
    If Not bDocDeServicio Then
        For Each itmX In lArticulo.ListItems    'Aca la mercadería siempre es toda a entregar
            ' (* -1)  ----> Porque le sumo al stock
            Cons = "Select * From Articulo Where ArtID = " & CLng(Mid(itmX.Key, 2, Len(itmX.Key))) _
                & " And ArtTipo = " & paTipoArticuloServicio
            Set RsCon = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If RsCon.EOF Then
                'No es del tipo de servicios
                RsCon.Close
                MarcoStockVenta CLng(tUsuario.Tag), CLng(Mid(itmX.Key, 2, Len(itmX.Key))), (CCur(itmX.Text) * -1), 0, 0, TipoDocumento.Contado, gDocumento, paCodigoDeSucursal
            Else
                RsCon.Close
            End If
        Next
    Else
        Cons = "Update Servicio Set SerDocumento = Null Where SerDocumento = " & gDocumento
        cBase.Execute Cons
    End If
    '----------------------------------------------------------------------------------

    aTexto = "Contado " & Trim(cSucursal.Text) & " " & Trim(RsAux!DocSerie) & " " & RsAux!DocNumero & " (" & Format(lFecha.Caption, "dd/mm/yy hh:mm") & ")"
    clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.AnulacionDeDocumentos, paCodigoDeTerminal, CLng(tUsuario.Tag), gDocumento, _
                                            Descripcion:=aTexto, Defensa:=Trim(cMotivos.Text), idCliente:=gCliente
        
    cBase.CommitTrans    'FIN DE TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Screen.MousePointer = 0
    
    LimpioFicha
    Foco cSucursal
    Exit Sub

errorBT:
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación."
    Screen.MousePointer = 0: Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación."
    Screen.MousePointer = 0
End Sub

Private Function ValidoAnularCredito() As Boolean
    
    ValidoAnularCredito = False
    
    If HayRemitos Then GoTo Salir
    If HayNota Then GoTo Salir
    If HayEnvio Then GoTo Salir
    If CreditoConRecibos() Then GoTo Salir
    Exit Function
    
Salir:
    ValidoAnularCredito = True
End Function

Private Sub AnuloCredito()
Dim RsCon As rdoResultset

    Screen.MousePointer = 11
    'Agregue condición de recibos, remitos, envíos.
    If ValidoAnularCredito() Then
        Screen.MousePointer = 0
        MsgBox aMensaje, vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    Dim sMsgErr As String
    
    
    On Error GoTo errorBT
    Screen.MousePointer = 11
    FechaDelServidor        'Saco la fecha del servidor
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
    Cons = "Select * from Documento Where DocCodigo = " & gDocumento
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux!DocFModificacion <> gFechaDocumento Then
        sMsgErr = "El documento ha sido modificado por otro usuario. Vuelva a cargar los datos."
'        Screen.MousePointer = 0
'        MsgBox "El documento ha sido modificado por otro usuario. Vuelva a cargar los datos.", vbExclamation, "ATENCIÓN"
        RsAux.Close
        'Fuerzo el error.
        RsAux.Edit
    End If
    
    'Actualizo el documento---------------------------------------------------
    RsAux.Edit
    RsAux!DocFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    RsAux!DocAnulado = True
    RsAux.Update
    '------------------------------------------------------------------------------
    
    If CreditoConRecibos Then
        sMsgErr = "El documento no se puede anular debido a que tiene asociado RECIBOS de PAGO"
        'Provoco error.
        RsAux.Close
        RsAux.Edit
    End If
    
    AnuloInstalacion
    
    'Agrego la Mercadería al Stock-------------------------------------------------
    If Not bDocDeServicio Then
        For Each itmX In lArticulo.ListItems
            Cons = "Select * From Articulo Where ArtID = " & CLng(Mid(itmX.Key, 2, Len(itmX.Key))) _
                & " And ArtTipo = " & paTipoArticuloServicio
            Set RsCon = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If RsCon.EOF Then
                'No es del tipo de servicios
                RsCon.Close
                ' (* -1)  ----> Porque le sumo al stock
                MarcoStockVenta CLng(tUsuario.Tag), CLng(Mid(itmX.Key, 2, Len(itmX.Key))), (CCur(itmX.Text) * -1), 0, 0, TipoDocumento.Credito, gDocumento, paCodigoDeSucursal
            Else
                RsCon.Close
            End If
        Next
    Else
        Cons = "Update Servicio Set SerDocumento = Null Where SerDocumento = " & gDocumento
        cBase.Execute Cons
    End If
    '-----------------------------------------------------------------------------------
    
    '-----------------------------------------------------------------------------------
    aTexto = "Crédito " & Trim(cSucursal.Text) & " " & Trim(RsAux!DocSerie) & " " & RsAux!DocNumero & " (" & Format(lFecha.Caption, "dd/mm/yy hh:mm") & ")"
    clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.AnulacionDeDocumentos, paCodigoDeTerminal, CLng(tUsuario.Tag), gDocumento, _
                                        Descripcion:=aTexto, Defensa:=Trim(cMotivos.Text), idCliente:=gCliente
    
    RsAux.Close
    cBase.CommitTrans    'FIN DE TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Screen.MousePointer = 0
    LimpioFicha
    Foco cSucursal
    Exit Sub
    
errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación."
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    If sMsgErr = "" Then sMsgErr = "No se ha podido realizar la transacción. Reintente la operación."
    clsGeneral.OcurrioError sMsgErr, Err.Description
End Sub

Private Sub AnuloReciboDePago()
txtERROR = ""
Dim idReciboAImprimir As Long

    idReciboAImprimir = 0
    Screen.MousePointer = 11
    Dim bNotaMultiple As Boolean, bNotaDeb As Boolean
    bNotaMultiple = ValidoReciboEnNotaMultiple(gDocumento)
    
    
    'Verifico si hay algun recibo que cubra una cuota superiror para algunos de los documentos---------------------
    For Each itmX In lRecibo.ListItems
        
        If CLng(Mid(itmX.Key, 2, Len(itmX.Key))) <> 40 Then 'And Not bNotaMultiple
            If TengoRecibosPosteriores(CLng(itmX.Tag), gDocumento) Then
                MsgBox "Documento Modificado: se cobraron cuotas posteriores a la seleccionada.", vbExclamation, "Datos Modificados"
                RsAux.Close: Screen.MousePointer = 0: Exit Sub
            End If
        ElseIf CLng(Mid(itmX.Key, 2, Len(itmX.Key))) = 40 And Not bNotaMultiple Then
            bNotaDeb = True
        End If
    Next
    '-----------------------------------------------------------------------------------------------------------------------------
    
    txtERROR = "02- F/modificación"
    Cons = "Select * from Documento Where DocCodigo = " & gDocumento
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
    If RsAux!DocFModificacion <> CDate(gFModificacionDocQAnulo) Then
        Screen.MousePointer = 0
        MsgBox "El documento ha sido modificado por otro usuario. Vuelva a cargar los datos.", vbExclamation, "ATENCIÓN"
        RsAux.Close
        Exit Sub
    End If
    Dim bVaANotaMultiple As Boolean
    bVaANotaMultiple = False
    If Format(RsAux!DocFecha, "dd/mm/yyyy") = Format(gFechaServidor, "dd/mm/yyyy") Then
        bVaANotaMultiple = ReciboConMoraANotaMultiple(gDocumento)
    End If
    On Error GoTo errorBT
    Screen.MousePointer = 11
    
    'Verifico que el documento no este asignado a alguna cuenta
    txtERROR = "03- Cta documento"
    If VerificoCuentaDocumento(gDocumento) Then
        MsgBox "El documento seleccionado está asignado a cuentas o colectivos." & Chr(vbKeyReturn) & "Para anularlo primero debe quitar la asignación.", vbInformation, "Documento Asignado a Cuentas"
        Screen.MousePointer = 0: Exit Sub
    End If

    FechaDelServidor        'Saco la fecha del servidor
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
    txtERROR = "04- Actualizando"
    '1) Actualizo el documento-----------------------------------------------------------------------------------
    RsAux.Edit
    RsAux!DocFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    RsAux.Update:
    RsAux.Close
    txtERROR = "06- Anulando"
    'idReciboAImprimir = AnuloReciboEnEfectivo((CCur(lImporte.Caption) < 0), bNotaMultiple Or bNotaDeb)
    Dim CAERecibo As clsCAEDocumento
    Dim DocCGSA As New clsDocumentoCGSA
    Set DocCGSA = AnuloReciboEnEfectivo((CCur(lImporte.Caption) < 0), bNotaMultiple Or bNotaDeb, CAERecibo)
    idReciboAImprimir = DocCGSA.Codigo
    
    'Si es nota deb múltiple o no es del día y tiene not. deb.
    If bNotaMultiple Or bNotaDeb Then
    
        Dim ivaND As Currency, totalND As Currency
        Dim docNotaDebNeg As Long, idNotaActual As Long
    
        If Not bNotaMultiple Then
            
            For Each itmX In lRecibo.ListItems
                If Val(Mid(itmX.Key, 2)) = TipoDocumento.NotaDebito Then
                    idNotaActual = itmX.Tag
                End If
            Next
            
            Cons = "SELECT DocTotal, DocIVA FROM Documento WHERE DocCodigo = " & idNotaActual & " AND DocTipo = 40"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            ivaND = RsAux("DocIVA")
            totalND = RsAux("DocTotal")
            RsAux.Close
        
        Else
            
            'Inserto una nota de débito en negativo por el monto de la mora.
            Cons = "SELECT NSFImporteTotal, NSFIva FROM NotasSinFacturar WHERE NSFIDRecibo = " & gDocumento
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
            'SI ES EOF PREFIERO QUE DE NULL y se analiza.
            ivaND = RsAux("NSFIva")
            totalND = RsAux("NSFImporteTotal")
            
            RsAux.Close
        End If
        docNotaDebNeg = GeneroNotaDebitoNegativa(totalND, ivaND, idReciboAImprimir, 1)
    ElseIf bVaANotaMultiple Then
        Cons = "UPDATE NotasSinFacturar SET NSFEstado = 9 Where NSFIDRecibo = " & gDocumento
        cBase.Execute Cons
    End If
    
    EmitirCFE DocCGSA, CAERecibo, paCodigoDGI
    
    txtERROR = "08- Updateo fechas facturas"
    'Updateo fecha de modificacion de las facturas orginiales
    Cons = "Update Documento Set DocFModificacion = GetDate() " & _
            " Where DocCodigo IN ( Select DPaDocASaldar from DocumentoPago Where DPaDocQSalda = " & gDocumento & ")"
    cBase.Execute Cons

    txtERROR = "09- Suceso"
    aTexto = "Recibo " & Trim(cSucursal.Text) & " " & Trim(tSerie.Text) & " " & Trim(tNumero.Text) & " (" & Format(lFecha.Caption, "dd/mm/yy hh:mm") & ")"
    Dim aFAnuladas As String: aFAnuladas = ""
    For Each itmX In lRecibo.ListItems: aFAnuladas = aFAnuladas & itmX.Text & ", ": Next
    aFAnuladas = "Facturas asociadas: " & Mid(aFAnuladas, 1, Len(aFAnuladas) - 2)
    aFAnuladas = aFAnuladas & vbCrLf & Trim(cMotivos.Text)
    clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.AnulacionDeDocumentos, paCodigoDeTerminal, CLng(tUsuario.Tag), gDocumento, _
                        Descripcion:=aTexto, Defensa:=aFAnuladas, idCliente:=gCliente
    
    cBase.CommitTrans    'FIN DE TRANSACCION-2-----------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    Dim sPaso As String
    Dim resM As VbMsgBoxResult
    If docNotaDebNeg > 0 Then
        sPaso = FirmoCFE(docNotaDebNeg)
        Do While sPaso <> ""
            resM = MsgBox("ATENCIÓN no se firmó el documento" & vbCrLf & sPaso & vbCrLf & vbCrLf & "Presione SI para reintentar" & vbCrLf & " Presione NO para abandonar ", vbExclamation + vbYesNo, "ATENCIÓN")
            If resM = vbNo Then Exit Do
            sPaso = FirmoCFE(docNotaDebNeg)
        Loop
    End If
    'If EmitirCFE(oDoc, CAE) <> "" Then RsAux.Close: RsAux.Edit
    'Imprimo si se emitio un recibo c/importe negativo
    'If idReciboAImprimir <> 0 Then ImprimoReciboPago idReciboAImprimir
    'If docNotaDebNeg <> 0 Then ImprimoReciboSeniaONotaDebito docNotaDebNeg, "", paDNDebito
    
    Screen.MousePointer = 0
    
    Dim FacturaAAnular As Long
    If gTipoPago = TipoPagoSolicitud.ChequeDiferido Then
        FacturaAAnular = CLng(lRecibo.ListItems(1).Tag)
    Else
        FacturaAAnular = 0
        If lRecibo.ListItems.Count = 1 Then
            FacturaAAnular = CLng(lRecibo.ListItems(1).Tag)
            'Consulto para ver si hay algun otro recibo de pago
            Cons = "Select * from DocumentoPago, Documento " _
                   & " Where DPaDocASaldar = " & FacturaAAnular _
                   & " And DPaDocQSalda = DocCodigo " _
                   & " And DocAnulado = 0"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux.EOF Then FacturaAAnular = 0    'Hay otros recibos que pagan la Factura
            RsAux.Close
        End If
    End If
    
    LimpioFicha
    
    If gTipoPago = TipoPagoSolicitud.ChequeDiferido Then
        MsgBox "Presione Aceptar, para que el sistema edite la factura asociada al recibo." & Chr(vbKeyReturn) & "Recuerde Anularla.", vbInformation, "ATENCIÓN"
        prmIDDocumento = 0
        CargoFacturaAAnular FacturaAAnular
        Foco tUsuario
    
    Else
        If FacturaAAnular <> 0 Then
            If MsgBox("Presione Aceptar, para que el sistema edite la factura asociada al recibo.", vbInformation + vbOKCancel, "ATENCIÓN") = vbOK Then
                prmIDDocumento = 0
                CargoFacturaAAnular FacturaAAnular
                Foco tUsuario
            Else
                Foco cSucursal
            End If
        Else
            Foco cSucursal
        End If
    End If
    Exit Sub

errorBT:
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción, reintente." & vbCrLf & txtERROR, Err.Description
    Screen.MousePointer = 0
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "No se ha podido realizar la transacción, reintente." & vbCrLf & txtERROR, Err.Description
    Screen.MousePointer = 0
End Sub

Private Function AnuloReciboEnEfectivo(EsReciboEnNegativo As Boolean, VaNotaNegativa As Boolean, CAERecibo As clsCAEDocumento) As clsDocumentoCGSA

Dim RsRec As rdoResultset, RsCuo As rdoResultset, RsCre As rdoResultset
Dim aCuota As Integer
Dim aProximoVto As Date, aPago As Date

Dim aRetorno As Long            'Id de recibo para imprimir x anulacion <> fecha
Dim bHaySaldo As Boolean
Dim auxMoraACuenta As Currency
        
    aRetorno = 0

    'Hay que volver atras el pago --> Actualizar la deuda
    For Each itmX In lRecibo.ListItems
        
        If Val(Mid(itmX.Key, 2)) <> TipoDocumento.NotaDebito Then
            auxMoraACuenta = 0
    
            aPago = CDate("01/01/1800")
            bHaySaldo = False
            If itmX.SubItems(1) = "E" Then aCuota = 0 Else: aCuota = CInt(itmX.SubItems(1))
            txtERROR = "06.1- Anulando"
            'Actualizo Tabla Credito--------------------------------------------------
            'VaCuota, UltimoPago, SaldoFactura, Mora, Cumplimiento, Puntaje, ProximoVto
            Cons = "Select * from Credito Where CreFactura = " & itmX.Tag
            Set RsCre = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
            txtERROR = "06.2- Anulando"
            'Actualizo Tabla CreditoCuota--------(Saldo, Mora, UltimoPago)------------------
            Cons = "Select * from CreditoCuota" _
                    & " Where CCuCredito = " & RsCre!CreCodigo _
                    & " And CCuNumero = " & aCuota
            Set RsCuo = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
            If Not IsNull(RsCuo!CCuSaldo) Then
                If RsCuo!CCuSaldo > 0 Then bHaySaldo = True
            End If
            
            'Si Valor-Saldo = Amortizacion  --> Hay un Solo Recibo (o no Amortiza el Saldo 5/2000 !!!)
            aProximoVto = RsCuo!CCuVencimiento
            If RsCuo!CCuValor - RsCuo!CCuSaldo = CCur(itmX.SubItems(2)) Then
                RsCuo.Edit
                RsCuo!CCuSaldo = RsCuo!CCuSaldo + CCur(itmX.SubItems(2))
                If Not IsNull(RsCuo!CCuMora) Then
                    RsCuo!CCuMora = RsCuo!CCuMora - CCur(itmX.SubItems(3))
                    
                    txtERROR = "06.3- Anulando"
                    'Debo verificar que mora pago a cuenta desde la ultima fecha de pago (si es que pago)---------------------
                    Cons = "Select * from DocumentoPago, Documento " _
                            & " Where DPaDocASaldar = " & itmX.Tag _
                            & " And DPaDocQSalda <> " & gDocumento _
                            & " And DPaDocQSalda = DocCodigo" _
                            & " And DPaCuota = " & aCuota & " And DocAnulado = 0" & " Order by DocFecha ASC"
                    Set RsRec = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    Do While Not RsRec.EOF
                        If RsRec!DocFecha > aPago And RsRec!DPaAmortizacion <> 0 Then
                            aPago = RsRec!DocFecha
                            auxMoraACuenta = 0
                        Else
                            If Not IsNull(RsRec!DPaMora) Then auxMoraACuenta = auxMoraACuenta + RsRec!DPaMora
                        End If
                        RsRec.MoveNext
                    Loop
                    RsRec.Close
                    '------------------------------------------------------------------------------------------------------------------------------------
                
                    RsCuo!CCuMoraACuenta = auxMoraACuenta
                End If
                
                RsCuo!CCuUltimoPago = Null
                RsCuo.Update
                
            Else    'Hay que sacar la fecha del pago del recibo para poner en ultimo pago
                txtERROR = "06.4- Anulando"
                'Debo verificar que mora pago a cuenta desde la ultima fecha de pago (si es que pago)---------------------
                Cons = "Select * from DocumentoPago, Documento " _
                        & " Where DPaDocASaldar = " & itmX.Tag _
                        & " And DPaDocQSalda <> " & gDocumento _
                        & " And DPaDocQSalda = DocCodigo" _
                        & " And DPaCuota = " & aCuota & " And DocAnulado = 0 Order by DocFecha ASC"
                Set RsRec = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                Do While Not RsRec.EOF
                    If RsRec!DocFecha > aPago And RsRec!DPaAmortizacion <> 0 Then
                        aPago = RsRec!DocFecha
                        auxMoraACuenta = 0
                    Else
                        If Not IsNull(RsRec!DPaMora) Then auxMoraACuenta = auxMoraACuenta + RsRec!DPaMora
                    End If
                    RsRec.MoveNext
                Loop
                RsRec.Close
                '------------------------------------------------------------------------------------------------------------------------------------
                
                RsCuo.Edit
                RsCuo!CCuSaldo = RsCuo!CCuSaldo + CCur(itmX.SubItems(2))
                If Not IsNull(RsCuo!CCuMora) Then RsCuo!CCuMora = RsCuo!CCuMora - CCur(itmX.SubItems(3))
                RsCuo!CCuMoraACuenta = auxMoraACuenta
                If aPago > CDate("01/01/1800") Then RsCuo!CCuUltimoPago = Format(aPago, "mm/dd/yyyy hh:mm:ss") Else RsCuo!CCuUltimoPago = Null
                RsCuo.Update
            End If
            RsCuo.Close
            txtERROR = "06.5- Anulando"
            'Edito el Resultset del Credito---------------------------------------
            Dim aVoyCuota As Integer
            RsCre.Edit
            
            If Not IsNull(RsCre!CreVaCuota) Then
                If Trim(RsCre!CreVaCuota) = "E" Then aVoyCuota = 0 Else: aVoyCuota = CInt(RsCre!CreVaCuota)
                If aVoyCuota = aCuota Then
                    aVoyCuota = aVoyCuota - 1
                    If aVoyCuota = 0 Then
                        'Veo si realmente esta almacenada la entrega
                        Cons = "Select * from CreditoCuota" _
                                & " Where CCuCredito = " & RsCre!CreCodigo _
                                & " And CCuNumero = 0"
                        Set RsCuo = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                        If Not RsCuo.EOF Then RsCre!CreVaCuota = "E" Else: RsCre!CreVaCuota = Null
                        RsCuo.Close
                        
                    Else
                        If aVoyCuota < 0 Then RsCre!CreVaCuota = Null Else: RsCre!CreVaCuota = CStr(aVoyCuota)
                    End If
                End If
                
                If Not bHaySaldo Then       'Si no quedaba saldo --> pago toda la cuota, actualizo el cumplimiento
                    If InStr(RsCre!CreCumplimiento, ".") <> 0 Then
                        RsCre!CreCumplimiento = Mid(RsCre!CreCumplimiento, 1, InStr(RsCre!CreCumplimiento, ".") - 2) & "." & Trim(Mid(RsCre!CreCumplimiento, InStr(RsCre!CreCumplimiento, "."), Len(RsCre!CreCumplimiento)))
                    Else
                        RsCre!CreCumplimiento = Mid(RsCre!CreCumplimiento, 1, Len(Trim(RsCre!CreCumplimiento)) - 1) & "."
                    End If
                End If
        
            End If
            Dim newSaldo As Currency
            
            If aPago = "01/01/1800" Then
                'Valido si hay pago de ctas anteriores --> para poner último pago al Credito (no a la cuota) 24/10/2002
                If Not IsNull(RsCre!CreVaCuota) Then
                    txtERROR = "06.6- Anulando"
                    Cons = "Select Top 1 *  from CreditoCuota" _
                            & " Where CCuCredito = " & RsCre!CreCodigo _
                            & " And CCuUltimoPago Is Not null " _
                            & " Order by CCuNumero desc"
                    Set RsCuo = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If Not RsCuo.EOF Then aPago = RsCuo!CCuUltimoPago
                    RsCuo.Close
                End If
            End If
            
            If aPago <> "01/01/1800" Then
                RsCre!CreUltimoPago = aPago
            Else
                RsCre!CreUltimoPago = Null
            End If
            ' ----------------------------------------------------------------------------------------- Fin del Cambio 24/10/2002
            
            If Not IsNull(RsCre!CreMora) Then RsCre!CreMora = RsCre!CreMora - CCur(itmX.SubItems(3))
            
            newSaldo = RsCre!CreSaldoFactura + CCur(itmX.SubItems(2))
            
            RsCre!CreSaldoFactura = newSaldo
            
            If newSaldo > 0 Then
                RsCre!CrePuntaje = Null
                RsCre!CreProximoVto = Format(aProximoVto, "mm/dd/yyyy hh:mm:ss")
            
                '25/10/2002 como el Va cuota es nulo --> estoy anulando un recibo que anulo el primer pago (por el crevacuota nulo)
                 ' Anulo un Rec en negativo por lo tanto queda valido el recibo anterior
                 'El 27/11/2002 agregue condicion para recibos en negativos
                 If aPago <> "01/01/1800" And IsNull(RsCre!CreVaCuota) And EsReciboEnNegativo Then
                    RsCre!CreCumplimiento = FormatoCumplimiento(0, aProximoVto, Trim(RsCre!CreCumplimiento), dFPago:=aPago)
                End If
                
            Else    'Anulo un recibo que es una nota sobre un recibo anulado en otro día 17/12/2001
                     'Aca en tra cuando éste Anulaba la Última Cuota    24/10/2002
                RsCre!CreProximoVto = Null
                RsCre!CreCumplimiento = FormatoCumplimiento(0, aProximoVto, Trim(RsCre!CreCumplimiento), dFPago:=aPago)
            End If
            RsCre.Update
            RsCre.Close
            
        End If
    Next
    txtERROR = "06.7- Anulando"
    'aRetorno = GeneroReciboPorAnulacion(False, VaNotaNegativa)
    
    Set AnuloReciboEnEfectivo = GeneroReciboPorAnulacion(CAERecibo, False, VaNotaNegativa)
    
End Function

Public Function CargoInfoCFE(ByVal idDoc As Long, ByRef tipoCFE As Byte) As clsClienteCFE

    tipoCFE = 0
    Dim oCliCFE As New clsClienteCFE
    'Intento levantar la información del ecomprobante.
    Dim rsCliente As rdoResultset
    
    Cons = "SET QUOTED_IDENTIFIER ON SET CONCAT_NULL_YIELDS_NULL ON SET ANSI_PADDING ON SET ANSI_WARNINGS ON SET ANSI_NULLS ON SET ARITHABORT ON SELECT EComTipo, EcomXml.value('(/*:CFE//*:Encabezado/*:Receptor/*:TipoDocRecep)[1]', 'tinyint') TipoDoc, " & _
            "EcomXml.value('(/*:CFE//*:Encabezado/*:Receptor/*:DocRecep)[1]', 'nvarchar(20)') Documento, " & _
            "EcomXml.value('(/*:CFE//*:Encabezado/*:Receptor/*:RznSocRecep)[1]', 'nvarchar(100)') Nombre, " & _
            "EcomXml.value('(/*:CFE//*:Encabezado/*:Receptor/*:DirRecep)[1]', 'nvarchar(100)') Direccion, " & _
            "EcomXml.value('(/*:CFE//*:Encabezado/*:Receptor/*:CiudadRecep)[1]', 'nvarchar(20)') Localidad, " & _
            "EcomXml.value('(/*:CFE//*:Encabezado/*:Receptor/*:DeptoRecep)[1]', 'nvarchar(20)') Departamento " & _
            "FROM eComprobantes WHERE EComID = " & idDoc
    Set rsCliente = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsCliente.EOF Then
        With oCliCFE
            .Codigo = oCliente.Codigo
            If Not IsNull(rsCliente("Nombre")) Then .NombreCliente = Trim(rsCliente!Nombre)
            If Not IsNull("EComTipo") Then tipoCFE = rsCliente("EComTipo")
            If Not IsNull(rsCliente("Direccion")) Then .Direccion.Domicilio = Trim(rsCliente("Direccion"))
            If Not IsNull(rsCliente("Departamento")) Then .Direccion.Departamento = Trim(rsCliente("Departamento"))
            If Not IsNull(rsCliente("Localidad")) Then .Direccion.Localidad = Trim(rsCliente("Localidad"))
            If Not IsNull(rsCliente("TipoDoc")) Then
                .CodigoDGICI = rsCliente("TipoDoc")
                If .CodigoDGICI = TD_RUT Then
                    .RUT = rsCliente("Documento")
                Else
                    .CI = rsCliente("Documento")
                End If
            Else
                .CodigoDGICI = TD_CI
            End If
        End With
    End If
    rsCliente.Close
    Set CargoInfoCFE = IIf(oCliCFE.Codigo > 0, oCliCFE, oCliente)
End Function

Private Function GeneroReciboPorAnulacion(ByRef CAEr As clsCAEDocumento, Optional ReciboDeSenia As Boolean = False, Optional VaNotaNegativa As Boolean = False) As clsDocumentoCGSA

    Dim aDocumentoRecibo As Long, Serie As String, Numero As Long
    'GeneroReciboPorAnulacion = 0
    
    aDocumentoRecibo = 0
    Dim tipoCFE As Byte
    Dim oDoc As New clsDocumentoCGSA
    Dim tipoCAE As Byte
    
    Dim oCli As clsClienteCFE
    Set oCli = CargoInfoCFE(gDocumento, tipoCFE)
    
    tipoCAE = IIf(tipoCFE = 111 Or oCli.RUT <> "", CFE_eFactura, CFE_eTicket)
    
    Set CAEr = New clsCAEDocumento
    Dim caeG As New clsCAEGenerador
    Set CAEr = caeG.ObtenerNumeroCAEDocumento(cBase, tipoCAE, paCodigoDGI)
    Set caeG = Nothing
    
    'Pido el Numero de Documento para hacer el RECIBO-------------
'    aTexto = NumeroDocumento(paDRecibo)
'    Serie = Trim(Mid(aTexto, 1, 1))
'    Numero = CLng(Trim(Mid(aTexto, 2, Len(aTexto))))

    'Inserto campos en la tabla documento---------------------------------------------------------------------
    Cons = "Select * from Documento Where DocCodigo = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.AddNew
    RsAux!DocFecha = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    RsAux!DocTipo = TipoDocumento.ReciboDePago
    RsAux!DocSerie = CAEr.Serie  'Serie
    RsAux!DocNumero = CAEr.Numero
    RsAux!DocCliente = gCliente
    RsAux!DocMoneda = Val(lMoneda.Tag)
    RsAux!DocTotal = CCur(lImporte.Caption) * -1
    RsAux!Dociva = 0             'IIf(VaNotaNegativa, 0, CCur(lImporte.Tag) * -1)               'El iva del recibo esta en el TAG de lImporte !!!
    RsAux!DocAnulado = 0
    RsAux!DocSucursal = paCodigoDeSucursal
    RsAux!DocUsuario = paCodigoDeUsuario
    RsAux!DocFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    RsAux.Update: RsAux.Close
    '---------------------------------------------------------------------------------------------------------------------
    
    Cons = "SELECT MAX(DocCodigo) From Documento" _
            & " WHERE DocTipo = " & TipoDocumento.ReciboDePago _
            & " AND DocSerie = '" & CAEr.Serie & "' AND DocNumero = " & CAEr.Numero
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly)
    aDocumentoRecibo = RsAux(0)
    RsAux.Close
    
    If Not ReciboDeSenia Then
        'RECIBO COMUN PAGO DE CUOTAS
        Dim aCuota As Integer
        For Each itmX In lRecibo.ListItems
            'Inserto relacion de Pagos --------------------------------------------------------------------------------------
            If UCase(Trim(itmX.SubItems(1))) = "E" Then aCuota = 0 Else aCuota = Val(itmX.SubItems(1))
            
            If Val(Mid(itmX.Key, 2)) <> TipoDocumento.NotaDebito Or Not VaNotaNegativa Then
                Cons = "Insert into DocumentoPago (DPaDocASaldar, DPaDocQSalda, DPaCuota, DPaDe, DPaAmortizacion, DPaMora)" _
                        & "Values (" & itmX.Tag & ", " _
                        & aDocumentoRecibo & ", " _
                        & aCuota & ", " _
                        & Val(itmX.SubItems(6)) & ", " _
                        & CCur(itmX.SubItems(2)) * -1 & ", " _
                        & CCur(itmX.SubItems(3)) * -1 & ")"
                cBase.Execute Cons
            End If
            '---------------------------------------------------------------------------------------------------------------------
        Next
    End If
    cBase.Execute "EXEC prg_PosInsertoDocumentosATickets '" & aDocumentoRecibo & "', " & oCnfgRecibo.ImpresoraTickets
    
    With oDoc
        Set .Cliente = oCli
        .Codigo = aDocumentoRecibo
        .Digitador = paCodigoDeUsuario
        .Emision = gFechaServidor
        .IVA = 0
        .Moneda.Codigo = 1
        .Numero = CAEr.Numero
        .Serie = CAEr.Serie
        .Sucursal = paCodigoDeSucursal
        .Tipo = TD_ReciboDePago
        .Total = CCur(lImporte.Caption)
    End With
    
    'EmitirCFE oDoc, CAEr, paCodigoDGI
    Set GeneroReciboPorAnulacion = oDoc
    
End Function

Private Function GeneroNotaDebitoNegativa(ByVal Total As Currency, ByVal IVA As Currency, ByVal idReciboViejo As Long, ByVal cuota As Integer) As Long

    Dim aDocNotaDebito As Long
    GeneroNotaDebitoNegativa = 0
    
    aDocNotaDebito = 0
        
    Dim tipoCAE As Byte
    tipoCAE = IIf(Val(lTitular.Tag) = 1, CFE_eFacturaNotaCredito, CFE_eTicketNotaCredito)
    
    Dim caeG As New clsCAEGenerador
    Dim CAE As New clsCAEDocumento
    Set CAE = caeG.ObtenerNumeroCAEDocumento(cBase, tipoCAE, paCodigoDGI, 0)
    Set caeG = Nothing
    
    Cons = "Select * from Documento Where DocCodigo = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.AddNew
    RsAux!DocFecha = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    RsAux!DocTipo = TipoDocumento.NotaCreditoMora
    RsAux!DocSerie = CAE.Serie
    RsAux!DocNumero = CAE.Numero
    RsAux!DocCliente = gCliente
    RsAux!DocMoneda = Val(lMoneda.Tag)
    RsAux!DocTotal = Total
    RsAux!Dociva = IVA
    RsAux!DocAnulado = 0
    RsAux!DocSucursal = paCodigoDeSucursal
    RsAux!DocUsuario = paCodigoDeUsuario
    RsAux!DocFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    RsAux.Update: RsAux.Close
    '---------------------------------------------------------------------------------------------------------------------
        
    Cons = "SELECT MAX(DocCodigo) From Documento" _
            & " WHERE DocTipo = " & TipoDocumento.NotaCreditoMora _
            & " AND DocSerie = '" & CAE.Serie & "' AND DocNumero = " & CAE.Numero
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly)
    aDocNotaDebito = RsAux(0)
    RsAux.Close
        
    Dim oDoc As New clsDocumentoCGSA
    With oDoc
        Set .Cliente = oCliente
        .Codigo = aDocNotaDebito
        .Digitador = paCodigoDeUsuario
        .Emision = gFechaServidor
        .IVA = Total - (Total / (1 + paIvaMora / 100))
        .Moneda.Codigo = Val(lMoneda.Tag)
        .Numero = CAE.Numero
        .Serie = CAE.Serie
        .Sucursal = paCodigoDeSucursal
        .Tipo = TipoDocumento.NotaCreditoMora
        .Total = Total
    End With
        
    Dim oRenglon As New clsDocumentoRenglon
    With oRenglon
        .Cantidad = 1
        .IVA = IVA
        .Precio = Total
        Set .Articulo = ProdInteresMora
    End With
    oDoc.Renglones.Add oRenglon
    
    Set oDoc.DocumentosAsociados = RetornoDocsRelacionados
    
    Dim moraT As Currency
    'moraT = Total * -1
    moraT = 0
    Cons = "Insert into DocumentoPago (DPaDocASaldar, DPaDocQSalda, DPaCuota, DPaDe, DPaAmortizacion, DPaMora)" _
                    & "Values (" & aDocNotaDebito & ", " _
                    & idReciboViejo & ", 1, 1, 0, " & moraT & ")"
    cBase.Execute Cons
    
    cBase.Execute "EXEC prg_PosInsertoDocumentosATickets '" & oDoc.Codigo & "', " & oCnfgRecibo.ImpresoraTickets
    GeneroNotaDebitoNegativa = aDocNotaDebito
        
End Function

Private Sub AnuloReciboDeSenia()

Dim sHacerSalidaDeCaja As Boolean
Dim idReciboAImprimir As Long            'Id de recibo para imprimir x anulacion <> fecha

    idReciboAImprimir = 0

    Screen.MousePointer = 11
    Cons = "Select * from Documento Where DocCodigo = " & gDocumento
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
    If RsAux!DocFModificacion <> gFechaDocumento Then
        Screen.MousePointer = 0
        MsgBox "El documento ha sido modificado por otro usuario. Vuelva a cargar los datos.", vbExclamation, "ATENCIÓN"
        RsAux.Close
        Exit Sub
    End If
    
    On Error GoTo errorBT
    Screen.MousePointer = 11
    FechaDelServidor        'Saco la fecha del servidor
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
    
    'Actualizo el documento--------------------------------------------------------------------------
    RsAux.Edit
    RsAux!DocFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    
 'Anule al emitir CFE
    'Veo si hago la salida de caja o no ---> Si la fecha es la mismo no se hace salida
'    If Format(RsAux!DocFecha, "dd/mm/yyyy") = Format(gFechaServidor, "dd/mm/yyyy") Then
'        sHacerSalidaDeCaja = False
'        RsAux!DocAnulado = True
'    Else
        sHacerSalidaDeCaja = True
'    End If
    
    RsAux.Update: RsAux.Close
    '---------------------------------------------------------------------------------------------------
    
    If sHacerSalidaDeCaja Then
                                
        'idReciboAImprimir = GeneroReciboPorAnulacion(ReciboDeSenia:=True)
        Dim DocCGSA As clsDocumentoCGSA
        Dim CAERecibo As clsCAEDocumento
        Set DocCGSA = GeneroReciboPorAnulacion(CAERecibo, ReciboDeSenia:=True)
        EmitirCFE DocCGSA, CAERecibo, paCodigoDGI
        'Debo desvincular el Recibo de la asignacion en la tabla cuenta documento
        Cons = "Delete CuentaDocumento Where CDoIdDocumento = " & gDocumento
        cBase.Execute Cons
        '---------------------------------------------------------------------------------------------------------------------
    End If
    
    aTexto = "Recibo " & Trim(cSucursal.Text) & " " & Trim(tSerie.Text) & " " & Trim(tNumero.Text) & " (" & Format(lFecha.Caption, "dd/mm/yy hh:mm") & ")"
    clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.AnulacionDeDocumentos, paCodigoDeTerminal, CLng(tUsuario.Tag), gDocumento, _
                    Descripcion:=aTexto, Defensa:="Recibo de Seña para Cuentas. " & Trim(cMotivos.Text), idCliente:=gCliente
    
    cBase.CommitTrans    'FIN DE TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    'Imprimo si se emitio un recibo c/importe negativo
    'If idReciboAImprimir <> 0 Then ImprimoReciboSeniaONotaDebito idReciboAImprimir, "Devolución de Seña", paDRecibo
    
    
    Screen.MousePointer = 0
    LimpioFicha
    Exit Sub

errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación."
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación."
End Sub


Private Sub AnuloNotaDevolucion()

    'HAY QUE ARREGLAR ESTA RUTINA PARA AGREGAR LOS ARTICULOS A LA FACTURA Y EL STOCK
       
    On Error GoTo errorBT
    Screen.MousePointer = 11
    FechaDelServidor        'Saco la fecha del servidor
    
    'Veo si el documento ha sido modificado (la Factura xq es Nota)---------------------------------------------------------
    Cons = "Select * from Documento Where DocCodigo = " & gFactura
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux!DocFModificacion <> gFechaDocumento Then
        Screen.MousePointer = 0
        MsgBox "La factura ha sido modificado por otro usuario. Vuelva a cargar los datos.", vbExclamation, "ATENCIÓN"
        RsAux.Close
        Exit Sub
    End If
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
    
    RsAux.Edit
    RsAux!DocFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    RsAux.Update
    RsAux.Close
    
    'Actualizo el documento (NOTA)---------------------------------------------------
    Cons = "Select * from Documento  Where DocCodigo = " & gDocumento
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.Edit
    RsAux!DocFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    RsAux!DocAnulado = True
    RsAux.Update
    RsAux.Close
    '--------------------------------------------------------------------------------------
    
    'Agrego la Mercadería al Stock-----------------------------------------------
    Dim aDefensa As String
    aDefensa = GraboStockXNota(gDocumento, gFactura, TipoDocumento.NotaDevolucion)
    aDefensa = aDefensa & vbCrLf & Trim(cMotivos.Text)
    
    aTexto = "Nota Devolución " & Trim(cSucursal.Text) & " " & Trim(tSerie.Text) & " " & Trim(tNumero.Text) & " (" & Format(lFecha.Caption, "dd/mm/yy hh:mm") & ")"
    clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.AnulacionDeDocumentos, paCodigoDeTerminal, CLng(tUsuario.Tag), gFactura, _
                                Descripcion:=aTexto, Defensa:=aDefensa, idCliente:=gCliente
    
    cBase.CommitTrans    'FIN DE TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    Screen.MousePointer = 0
    LimpioFicha
    Foco cSucursal
    Exit Sub
    
errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación."
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación."
End Sub

Private Sub AnuloNotaEspecial()

    On Error GoTo errorBT
    Screen.MousePointer = 11
    FechaDelServidor        'Saco la fecha del servidor
    ' *** CREO QUE FALTA MOVIMIENTO DE CAJA ***
    
    'Veo si el documento ha sido modificado (la Factura xq es Nota)---------------------------------------------------------
    If gFactura > 0 Then
        Cons = "Select * from Documento Where DocCodigo = " & gFactura
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux!DocFModificacion <> gFechaDocumento Then
            Screen.MousePointer = 0
            MsgBox "La factura ha sido modificado por otro usuario. Vuelva a cargar los datos.", vbExclamation, "ATENCIÓN"
            RsAux.Close: Exit Sub
        End If
    End If
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
    
    If gFactura > 0 Then
        RsAux.Edit
        RsAux!DocFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
        RsAux.Update: RsAux.Close
    End If
    
    'Actualizo el documento (NOTA)---------------------------------------------------
    Cons = "Select * from Documento  Where DocCodigo = " & gDocumento
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.Edit
    RsAux!DocFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    RsAux!DocAnulado = True
    RsAux.Update: RsAux.Close
    '--------------------------------------------------------------------------------------
    
    'Agrego la Mercadería al Stock-----------------------------------------------
    Dim aDefensa As String
    aDefensa = GraboStockXNota(gDocumento, gFactura, TipoDocumento.NotaEspecial)
    aDefensa = aDefensa & vbCrLf & Trim(cMotivos.Text)
    
    aTexto = "Nota Especial " & Trim(cSucursal.Text) & " " & Trim(tSerie.Text) & " " & Trim(tNumero.Text) & " (" & Format(lFecha.Caption, "dd/mm/yy hh:mm") & ")"
    clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.AnulacionDeDocumentos, paCodigoDeTerminal, CLng(tUsuario.Tag), gFactura, _
                    Descripcion:=aTexto, Defensa:=aDefensa, idCliente:=gCliente
    
    cBase.CommitTrans    'FIN DE TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    Screen.MousePointer = 0
    LimpioFicha
    Foco cSucursal
    Exit Sub
    
errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación."
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación."
End Sub

Private Sub AnuloNotaCreditoEnvio()

Dim aCredito As Long
Dim aSalidaDeCaja As Currency   'Salida de Caja de la Nota
Dim aMontoNota As Currency      'Monto a reintegrar como pago desde la nota

    On Error GoTo errProceso
    
    Dim mIDDisponibilidad As Long
    mIDDisponibilidad = ValidoDisponibilidad
    If mIDDisponibilidad = 0 Then Exit Sub
    
    Screen.MousePointer = 11
    aMensaje = ""
    
    'Solo anula notas de creditos emitidas por el envio
    '1)  Hace mov de caja para contrarestar y anula la nota, no reintegra importes a las cuotas.
    
    'Saco los Importes de Dovolucion y Salida de Caja------------------------------------------
    Cons = "Select * from Nota Where NotNota = " & gDocumento & " And NotFactura = " & gFactura
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not IsNull(RsAux!NotDevuelve) Then aMontoNota = RsAux!NotDevuelve Else: aMontoNota = 0
    If Not IsNull(RsAux!NotSalidaCaja) Then aSalidaDeCaja = RsAux!NotSalidaCaja Else: aSalidaDeCaja = 0
    RsAux.Close
    '---------------------------------------------------------------------------------------------------
    
    'Saco el Codigo del Credito-------------------------------------------------------------------
    Cons = "Select * from Credito Where CreFactura = " & gFactura
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    aCredito = RsAux!CreCodigo
    RsAux.Close
    '------------------------------------------------------------------------------------------------
    
    On Error GoTo errorBT
    FechaDelServidor        'Saco la fecha del servidor
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
    
    'Veo si el documento ha sido modificado (la Factura xq es Nota)---------------------------------------------------------
    Cons = "Select * from Documento Where DocCodigo = " & gFactura
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux!DocFModificacion <> gFechaDocumento Then
        Screen.MousePointer = 0
        aMensaje = "La factura ha sido modificado por otro usuario. Vuelva a cargar los datos."
        RsAux.Close
        GoTo errorET
        Exit Sub
    End If
    RsAux.Edit
    RsAux!DocFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    RsAux.Update
    RsAux.Close
    '-------------------------------------------------------------------------------------------------------------------------------
    
    'Actualizo el documento (Nota)--------------------------------------------------------------------
    Cons = "Select * from Documento Where DocCodigo = " & gDocumento
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.Edit
    RsAux!DocFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    RsAux!DocAnulado = True
    RsAux.Update
    RsAux.Close
    '-------------------------------------------------------------------------------------------------------------------------------

    'Movimiento de Caja------------------------
    MovimientoDeCaja paMCAnulacion, gFechaServidor, mIDDisponibilidad, CLng(lMoneda.Tag), aSalidaDeCaja, _
                                Trim(Trim(cDocumento.Text) & " " & cSucursal.Text & " " & tSerie.Text & tNumero.Text), False
        
    
    aTexto = "Nota Crédito " & Trim(cSucursal.Text) & " " & Trim(tSerie.Text) & " " & Trim(tNumero.Text) & " (" & Format(lFecha.Caption, "dd/mm/yy hh:mm") & ")"
    clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.AnulacionDeDocumentos, paCodigoDeTerminal, CLng(tUsuario.Tag), gFactura, _
                        Descripcion:=aTexto, Defensa:="Nota sobre Flete de envíos. " & Trim(cMotivos.Text), idCliente:=gCliente
    
    cBase.CommitTrans    'FIN DE TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Screen.MousePointer = 0
    LimpioFicha
    Foco cSucursal
    Exit Sub

errProceso:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al procesar los datos"
    Exit Sub
errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación."
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    If aMensaje = "" Then aMensaje = "No se ha podido realizar la transacción. Reintente la operación."
    clsGeneral.OcurrioError aMensaje
End Sub

Private Sub AnuloNotaCredito()

Dim aCredito As Long
Dim aSalidaDeCaja As Currency   'Salida de Caja de la Nota
Dim aMontoNota As Currency      'Monto a reintegrar como pago desde la nota

Dim aTotalF As Currency, aPrecio As Currency
Dim Fletes As String

Dim aCuotaOriginal As Currency, aEntregaOriginal As Currency        'Valores $$
Dim aCCuotas As Integer         'Cantidad de Cuotas (Sin Entrega)
Dim sConEntrega As Boolean
Dim aPorcentajePago As Currency
    
    On Error GoTo errProceso
    
    Dim mIDDisponibilidad As Long
    mIDDisponibilidad = ValidoDisponibilidad
    If mIDDisponibilidad = 0 Then Exit Sub
    
    Screen.MousePointer = 11
    aMensaje = ""
    
    'Saco el formato del reondeo desde la tabla moneda --------------------------------------------------
    Dim mPRound As String: mPRound = "1"
    Cons = "Select * from Moneda Where MonCodigo = " & Val(lMoneda.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!MonRedondeo) Then
            If Trim(RsAux!MonRedondeo) <> "" Then mPRound = Trim(RsAux!MonRedondeo)
        End If
    End If
    RsAux.Close
    '---------------------------------------------------------------------------------------------------------------
    
    'Como la Nota se Puede Anular si no se han pagado más cuotas, Hay  que reintegrar el monto a las cuotas
    'que quedan y Hacer una entrada de caja para la cuota parte de los articulos en las cuotas pagas.
    Fletes = CargoArticulosDeFlete
    
    'Saco los Importes de Dovolucion y Salida de Caja------------------------------------------
    Cons = "Select * from Nota Where NotNota = " & gDocumento & " And NotFactura = " & gFactura
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not IsNull(RsAux!NotDevuelve) Then aMontoNota = RsAux!NotDevuelve Else: aMontoNota = 0
    If Not IsNull(RsAux!NotSalidaCaja) Then aSalidaDeCaja = RsAux!NotSalidaCaja Else: aSalidaDeCaja = 0
    RsAux.Close
    '---------------------------------------------------------------------------------------------------
    
    'Saco El Importe Total de la Factura (Sin los Articulos del Envio)-----------------------
    aTotalF = 0
    Cons = " Select * from Renglon Where RenDocumento = " & gFactura
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        If InStr(Fletes, RsAux!RenArticulo & ",") = 0 Then aTotalF = aTotalF + (RsAux!RenCantidad * RsAux!RenPrecio)
        RsAux.MoveNext
    Loop
    RsAux.Close
    '------------------------------------------------------------------------------------------------
    
    'Saco el Codigo del Credito-------------------------------------------------------------------
    Cons = "Select * from Credito Where CreFactura = " & gFactura
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    aCredito = RsAux!CreCodigo
    aCuotaOriginal = RsAux!CreValorCuota
    RsAux.Close
    '------------------------------------------------------------------------------------------------
    
    'Saco la Cantidad de Cuotas del Credito-----------------------------------------------------
    Cons = "Select * from CreditoCuota Where CCuCredito = " & aCredito
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    aCCuotas = 0: sConEntrega = False
    Do While Not RsAux.EOF
        aCCuotas = aCCuotas + 1
        If RsAux!CCuNumero = 0 Then sConEntrega = True: aCCuotas = aCCuotas - 1
        RsAux.MoveNext
    Loop
    RsAux.Close
    '------------------------------------------------------------------------------------------------
    
    'Valor de la Entrega
    If sConEntrega Then aEntregaOriginal = aTotalF - (aCuotaOriginal * aCCuotas)
    
    '% de la Nota sobre el Total Factura
    aPorcentajePago = (CCur(lImporte.Caption) * 100) / aTotalF / 100

    On Error GoTo errorBT
    FechaDelServidor        'Saco la fecha del servidor
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
    
    'Veo si el documento ha sido modificado (la Factura xq es Nota)---------------------------------------------------------
    Cons = "Select * from Documento Where DocCodigo = " & gFactura
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux!DocFModificacion <> gFechaDocumento Then
        Screen.MousePointer = 0
        aMensaje = "La factura ha sido modificado por otro usuario. Vuelva a cargar los datos."
        RsAux.Close
        GoTo errorET
        Exit Sub
    End If
    RsAux.Edit
    RsAux!DocFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    RsAux.Update
    RsAux.Close
    '-------------------------------------------------------------------------------------------------------------------------------
    
    'Actualizo el documento (Nota)--------------------------------------------------------------------
    Cons = "Select * from Documento Where DocCodigo = " & gDocumento
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.Edit
    RsAux!DocFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    RsAux!DocAnulado = True
    RsAux.Update
    RsAux.Close
    '-------------------------------------------------------------------------------------------------------------------------------

    GraboAjusteCuotasCredito aCredito, aCuotaOriginal, aEntregaOriginal, aMontoNota, aPorcentajePago, mPRound
    
    'Movimiento de Caja------------------------
    MovimientoDeCaja paMCAnulacion, gFechaServidor, mIDDisponibilidad, CLng(lMoneda.Tag), aSalidaDeCaja, _
                                Trim(Trim(cDocumento.Text) & " " & cSucursal.Text & " " & tSerie.Text & tNumero.Text), False
        
    
    'Actualizo la Mercadería al Stock-----------------------------------------------
    Dim aDefensa As String
    aDefensa = GraboStockXNota(gDocumento, gFactura, TipoDocumento.NotaCredito)
    aDefensa = aDefensa & vbCrLf & Trim(cMotivos.Text)
    
    aTexto = "Nota Crédito " & Trim(cSucursal.Text) & " " & Trim(tSerie.Text) & " " & Trim(tNumero.Text) & " (" & Format(lFecha.Caption, "dd/mm/yy hh:mm") & ")"
    clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.AnulacionDeDocumentos, paCodigoDeTerminal, CLng(tUsuario.Tag), gFactura, _
                                        Descripcion:=aTexto, Defensa:=aDefensa, idCliente:=gCliente
    
    cBase.CommitTrans    'FIN DE TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Screen.MousePointer = 0
    LimpioFicha
    Foco cSucursal
    Exit Sub

errProceso:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al procesar los datos"
    Exit Sub
errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación."
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    If aMensaje = "" Then aMensaje = "No se ha podido realizar la transacción. Reintente la operación."
    clsGeneral.OcurrioError aMensaje
End Sub


Private Sub GraboAjusteCuotasCredito(Credito As Long, cuota As Currency, Entrega As Currency, MontoNota As Currency, PorcentajePago As Currency, mPatronRnd As String)

Dim aCuotaParte As Currency     'Cuota Parte del Articulo
Dim aSaldoCredito As Currency   'Nuevo saldo del Credito
Dim aSaldoCuota As Currency     'Nuevo saldo de la cuota
    
    Cons = "Select * from CreditoCuota Where CCuCredito = " & Credito
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        'Saco el Valor a Reintegrar en cada Cuota-------------------------------------------------------------------------------
        If RsAux!CCuNumero <> 0 Then
            aCuotaParte = Redondeo(cuota * PorcentajePago, mPatronRnd)          'Aplico el % al Valor de la Cuota
        Else
            aCuotaParte = Redondeo(Entrega * PorcentajePago, mPatronRnd)        'Aplico el % al Valor de la Cuota
        End If
        '------------------------------------------------------------------------------------------------------------------------------
        
        'Veo cuanto es lo que reintegro del pago de la cuota-------------------
        If MontoNota >= aCuotaParte Then
            MontoNota = MontoNota - aCuotaParte
            aSaldoCuota = 0
        Else
            aSaldoCuota = RsAux!CCuSaldo + (aCuotaParte - MontoNota)
            MontoNota = 0
        End If
        '--------------------------------------------------------------------------------
        
        Cons = "Update CreditoCuota " _
                & " Set CCuValor = CCuValor + " & aCuotaParte & ", " _
                & " CCuSaldo = " & aSaldoCuota _
                & " Where CCuCredito = " & Credito _
                & " And CCuNumero = " & RsAux!CCuNumero
        
        cBase.Execute Cons

        RsAux.MoveNext
    Loop
    RsAux.Close
    
    'Actualizo la TablaCredito CreSaldoFactura----------------------------------------------------------------
    Cons = "Select Sum(CCuSaldo) from CreditoCuota Where CCuCredito = " & Credito
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    aSaldoCredito = RsAux(0)
    RsAux.Close
    
    Cons = "Update Credito Set CreSaldoFactura = " & aSaldoCredito _
           & " Where CreCodigo = " & Credito
    cBase.Execute Cons
    
End Sub

Private Function GraboStockXNota(Nota As Long, Factura As Long, TipoDoc As Integer) As String
Dim rs2 As rdoResultset
Dim aTxt As String
    'Ver Caso: Nota q salio con retiro y entro mercaderia al local --> +1 loc pero se anula la nota (habria que anular dev)
    aTxt = ""
    Cons = "Select * from Renglon, Articulo " & _
               " Where RenDocumento = " & Nota & _
               " And RenArticulo = ArtID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        
        If RsAux!RenARetirar > 0 Then
            Cons = "Update Renglon Set RenARetirar = RenARetirar + " & RsAux!RenARetirar _
                    & " Where RenDocumento = " & Factura _
                    & " And RenArticulo = " & RsAux!RenArticulo
            cBase.Execute (Cons)
            
            If RsAux!ArtTipo <> paTipoArticuloServicio Then
            
                'Cambio - ahora no ahcemos mov. fisicos
                MarcoStockXDevolucion RsAux!RenArticulo, RsAux!RenARetirar, RsAux!RenARetirar, _
                                              TipoLocal.Deposito, paCodigoDeSucursal, CLng(tUsuario.Tag), TipoDoc, Nota, True
            
                aTxt = aTxt & RsAux!RenARetirar & " " & Format(RsAux!ArtCodigo, "(#,000,000)") & " A Retirar (de Nota a Factura); "
                
            End If
        End If
        
        '- 08/08/2001 - saco rel. nota.
        Cons = "Select * From Devolucion " & _
                   " Where DevNota = " & Nota & _
                   " And DevArticulo = " & RsAux!RenArticulo & " And DevLocal is not Null"
                        
        Set rs2 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rs2.EOF Then
            aTxt = aTxt & "Quito " & rs2!DevCantidad & " " & Format(RsAux!ArtCodigo, "(#,000,000)") & " en F.Dev.; "
            
            rs2.Edit
            rs2!DevNota = Null
            rs2.Update
        End If
        rs2.Close
                                                  
        RsAux.MoveNext
    Loop
    RsAux.Close
        
    'Hay que ver que pasa con los registros en la Tabla Devolucion
    'Cons = "Delete Devolucion Where DevNota = " & Nota & _
                " And DevLocal is Null"
    Cons = "Update Devolucion Set DevAnulada = '" & Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss") & "'" & _
                " Where DevNota = " & Nota
    cBase.Execute Cons  'actualizo f anulada
    
    If Len(aTxt) > 2 Then aTxt = Mid(aTxt, 1, Len(aTxt) - 2)
    GraboStockXNota = aTxt
    
End Function

Private Function HayRemitos() As Boolean

    HayRemitos = False
    Cons = "SELECT Rtrim(DocSerie) + '-' + Convert(VarChar(6), DocNumero) FROM RemitoDocumento INNER JOIN DOCUMENTO ON RDoRemito = DocCodigo " & _
            " WHERE RDoDocumento = " & gDocumento & " AND DocAnulado = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        aMensaje = "El documento no se puede anular debido a que tiene asociado el REMITO Nº " & RsAux(0)
        HayRemitos = True
    End If
    RsAux.Close
    
End Function

Private Function HayNota() As Boolean
    
    'Valido existencia de Nota-----------------------------------------------------------------------------------------------
    HayNota = False
    Cons = "Select * from Nota, Documento " _
            & " Where NotFactura = " & gDocumento _
            & " And NotNota = DocCodigo" _
            & " And DocAnulado = 0"     '0 = Falso
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        aMensaje = "El documento no se puede anular debido a que tiene asociado la NOTA Nº " & Trim(RsAux!DocSerie) & Trim(RsAux!DocNumero)
        HayNota = True
    End If
    RsAux.Close
    
End Function

Private Function HayEnvioCobroFlete(ByVal bMensaje As Boolean) As Boolean

    Cons = "Select * from Envio " _
            & " Where EnvTipo = " & TipoEnvio.Entrega _
            & " And EnvDocumentoFactura = " & gDocumento & " AND EnvDocumento <> EnvDocumentoFactura AND EnvEstado <> 4"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        If bMensaje Then
            MsgBox "ATENCIÓN!!!, El documento paga el flete del ENVÍO Nº " & RsAux!EnvCodigo & ", al anularlo el sistema cambiará la forma de pago de dicho envío a " & paFPagoAnulaDocumentoNombre & ".", vbInformation, "ATENCIÓN (prm:FormaPagoEnvioDocAnulado)"
        End If
        HayEnvioCobroFlete = True
    End If
    RsAux.Close

End Function

Private Function HayEnvio() As Boolean
    
    'Valido existencia de Envio-----------------------------------------------------------------------------------------------
    HayEnvio = False
    Cons = "Select * from Envio " _
            & " Where EnvTipo = " & TipoEnvio.Entrega _
            & " And EnvDocumento = " & gDocumento
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not RsAux.EOF Then
        aMensaje = "El documento no se puede anular debido a que tiene asociado el ENVÍO Nº " & RsAux!EnvCodigo
        HayEnvio = True
    End If
    RsAux.Close
            
End Function

Private Sub HayInstalacion()
    
    Cons = "Select InsID from Instalacion " _
            & " Where InsTipoDocumento = 1 " _
            & " And InsDocumento = " & gDocumento
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not RsAux.EOF Then
        MsgBox "El documento tiene asociadas instalaciones sí lo anula se anulará la instalación.", vbInformation, "Instalaciones"
    End If
    RsAux.Close
            
End Sub



Private Function HayRecibosParaFactura(MayorA As Date, Factura As Long) As Boolean
    
    'Valido existencia de recibos con fechas posteriores a las de la nota
    HayRecibosParaFactura = False
    Cons = "Select * from DocumentoPago, Documento" _
            & " Where DPaDocQSalda = " & Factura _
            & " And DPaDocASaldar = DocCodigo" _
            & " And DocAnulado = 0" _
            & " And DocFecha > '" & Format(MayorA, "mm/dd/yyyy hh:mm:ss") & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not RsAux.EOF Then
        aMensaje = "El documento no se puede anular debido a que hay recibos de pago posteriores a la fecha de emisión de la nota."
        HayRecibosParaFactura = True
    End If
    RsAux.Close
            
End Function

Private Function HayNotasParaFactura(MayorA As Date, Factura As Long) As Boolean
    
    'Valido existencia de Nota-----------------------------------------------------------------------------------------------
    HayNotasParaFactura = False
    Cons = "Select * from Nota, Documento " _
            & " Where NotFactura = " & Factura _
            & " And NotNota = DocCodigo" _
            & " And DocAnulado = 0" _
            & " And DocFecha > '" & Format(MayorA, "mm/dd/yyyy hh:mm:ss") & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        aMensaje = "El documento no se puede anular debido a que hay devoluciones posteriores a la fecha de emisión de la nota seleccionada."
        HayNotasParaFactura = True
    End If
    RsAux.Close
    
End Function

'----------------------------------------------------------------------------------------------------------------------------
'   Anulacion de Recibos Y Anulacion de Notas (Las notas no se pueden anular si Pago Ch.Dif).
'   Esta rutina valida si la factura asociada fue paga con Cheques, si es así y no hay ningún cheque que haya
'   sido depositado se puede anular.  Si hay algun cheque depositado no se puede anular.
'----------------------------------------------------------------------------------------------------------------------------
Private Function FacturaPagaConCheques(Factura As Long, Optional Recibo As Boolean = False, Optional Nota As Boolean = False) As Boolean
    
    'Valido si la factura fue paga con cheque diferido----------------------------------------------------------------------------
    FacturaPagaConCheques = False
    Cons = "Select * from Documento, Credito " _
            & " Where DocCodigo = " & Factura _
            & " And DocCodigo = CreFactura"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        gTipoPago = RsAux!CreFormaPago
        If RsAux!CreFormaPago = TipoPagoSolicitud.ChequeDiferido Then
            If Recibo Then      'ANULANDO RECIBO
                'Fue Pago con cheques --->   verifico si alguno fue depositado
                Dim RsCD As rdoResultset
                'Cons = "Select * from ChequeDiferido" _
                    & " Where CDiDocumento = " & gDocumento _
                    & " And CDiDepositado = 1"
                'Set RsCD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                
                'If Not RsCD.EOF Then
                '    aMensaje = "El documento no se puede anular, la factura asociada  fue paga con cheques diferidos y algunos de los cheque está depositado."
                '    FacturaPagaConCheques = True
                'Else
                    MsgBox "El recibo fue pago con cheques diferidos, si lo anula, controle la factura y los cheques.", vbInformation, "ATENCIÓN"
                'End If
                'RsCD.Close
            End If
            
            If Nota Then        'Anulando NOTA
                aMensaje = "El documento no se puede anular, la factura asociada  fue paga con cheques diferidos."
                FacturaPagaConCheques = True
            End If
            
        End If
    End If
    RsAux.Close
    
End Function

Private Sub tUsuario_Change()
    tUsuario.Tag = 0
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And IsNumeric(tUsuario.Text) Then
        tUsuario.Tag = z_BuscoUsuarioDigito(Val(tUsuario.Text), Codigo:=True)
        If Val(tUsuario.Tag) = 0 Then tUsuario.Text = "": Exit Sub
        
        If Val(tUsuario.Tag) <> 0 Then
            If bAnular.Enabled Then bAnular.SetFocus Else: bCancelar.SetFocus
        End If
    End If
    
End Sub

Private Sub HabilitoCampos(Estado As Boolean)

    bCancelar.Enabled = Estado
    bAnular.Enabled = Estado
    MnuCancelar.Enabled = Estado
    MnuGrabar.Enabled = Estado
    
    tUsuario.Enabled = Estado
    cMotivos.Enabled = Estado
        
    If Estado Then
        tUsuario.BackColor = Obligatorio
        cMotivos.BackColor = Colores.Obligatorio
    Else
        tUsuario.BackColor = Inactivo
        cMotivos.BackColor = Inactivo
    End If
    
    Screen.MousePointer = 0
End Sub


'-----------------------------------------------------------------------------------------------------------------------
'   Este procedimiento se u¡invoca cuando se anula un recibo con cheques---> para que anule la factura
'---------------------------------------------------------------------------------------------------------------------
Private Sub CargoFacturaAAnular(Codigo As Long)

    On Error GoTo errCargar
    gFactura = 0
    prmPuedoAnular = True
    prmDocAnulado = False
    
    Cons = " Select * From Documento Left Outer Join Usuario On DocUsuario = UsuCodigo" & _
                                                    " Left Outer Join Moneda On DocMoneda = MonCodigo" & _
               " Where DocCodigo = " & Codigo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    gDocumento = Codigo
    
    If Not RsAux.EOF Then
        BuscoCodigoEnCombo cSucursal, RsAux!DocSucursal
        BuscoCodigoEnCombo cDocumento, TipoDocumento.Credito
        tSerie.Text = Trim(RsAux!DocSerie)
        tNumero.Text = RsAux!DocNumero
        
        lArticulo.ZOrder 0
        lTitulo.Caption = " ARTÍCULOS"
        lTitulo.Tag = "A"
        
        gCliente = RsAux!DocCliente
        gFechaDocumento = RsAux!DocFModificacion
        
        lFecha.Caption = Format(RsAux!DocFecha, "d-Mmm-yy hh:mm:ss")
        lFecha.Tag = RsAux!DocFecha
        If Not IsNull(RsAux!UsuIdentificacion) Then lUsuario.Caption = Trim(RsAux!UsuIdentificacion)
    
        lComentario.Caption = "N/D"
        If Not IsNull(RsAux!DocComentario) Then lComentario.Caption = Trim(RsAux!DocComentario)
    
        If Not IsNull(RsAux!MonSigno) Then lMoneda.Caption = Trim(RsAux!MonSigno)
        lMoneda.Tag = RsAux!DocMoneda
        lImporte.Caption = Format(RsAux!DocTotal, FormatoMonedaP)
        
        If RsAux!DocAnulado Then
            MsgBox "El documento seleccionado ya ha sido anulado.", vbInformation, "ATENCIÓN"
            prmPuedoAnular = False
            prmDocAnulado = True
        End If
               
    Else
        MsgBox "No existe un documento para las características ingresadas.", vbInformation, "ATENCIÓN"
        prmPuedoAnular = False
    End If
    RsAux.Close
    
    If Not prmPuedoAnular Then HabilitoCampos False: Exit Sub
    
    CargoCliente gCliente
    CargoArticulos gDocumento
    
    If prmDocAnulado Then HabilitoCampos False: Exit Sub
    
    If Format(gFechaDocumento, "yyyy/mm/dd") <> Format(gFechaServidor, "yyyy/mm/dd") Then
        MsgBox "El documento no se puede anular. Sólo se podrán anular los documentos del día.", vbInformation, "ATENCIÓN"
        HabilitoCampos False: Exit Sub
    End If
    
    'If prmPuedoAnular Then     'Veo la causa de no poder anular------------------------------------------------------
    If HayRemitos Then GoTo Salir
    If HayNota Then GoTo Salir
    If HayEnvio Then GoTo Salir

    If Not prmPuedoAnular And bQARetirarDif Then
        '1) Veo si el doc es de servicio
        DocumentoDeServicios gDocumento
        If bDocDeServicio Then prmPuedoAnular = True
        If Not prmPuedoAnular Then
            aMensaje = "El documento no se puede anular, las cantidades de venta no coinciden con las a retirar."
            GoTo Salir
        End If
    End If
    
    'If Not prmPuedoAnular Then
    '    aMensaje = "El documento no se puede anular, las cantidades de venta no coinciden con las a retirar."
    '    GoTo Salir
    'End If
    
    If prmPuedoAnular Then HabilitoCampos True
    Exit Sub
    
errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del documento."
    HabilitoCampos False: Exit Sub
Salir:
    Screen.MousePointer = 0
    MsgBox aMensaje, vbInformation, "ATENCIÓN"
    HabilitoCampos False
End Sub

'Private Sub ImprimoReciboPago(aRecibo)
'
'Dim JobSRep1 As Integer, JobSRep2 As Integer, Result As Integer
'Dim NombreFormula As String
'
'    On Error GoTo ErrCrystal
'    If Not InicializoCrystalEngine("Recibo.RPT") Then Screen.MousePointer = 0: Exit Sub
'
'    Screen.MousePointer = 11
'
'    'Cargo Propiedades para el reporte Contado --------------------------------
'    For i = 0 To CantForm - 1
'        NombreFormula = crObtengoNombreFormula(jobnum, i)
'
'        Select Case LCase(NombreFormula)
'            Case "": GoTo ErrCrystal
'            Case "nombredocumento": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & paDRecibo & "'")
'
'            Case "cliente": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lTitular.Caption) & "'")
'            Case "cedula":
'                If Val(lCiRuc.Tag) = 1 Then
'                    Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lCiRuc.Caption) & "'")
'                Else
'                    Result = crSeteoFormula(jobnum%, NombreFormula, "''")
'                End If
'
'            Case "ruc":
'                If Val(lCiRuc.Tag) = TipoCliente.Empresa Then
'                    Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lCiRuc.Caption) & "'")
'                Else
'                    Result = crSeteoFormula(jobnum%, NombreFormula, "''")
'                End If
'
'            Case "signomoneda": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lMoneda.Caption) & "'")
'            Case "nombremoneda": Result = crSeteoFormula(jobnum%, NombreFormula, "'(" & BuscoNombreMoneda(lMoneda.Tag) & ")'")
'
'            Case "usuario": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & BuscoDigitoUsuario(paCodigoDeUsuario) & "'")
'
'            Case Else: Result = 1
'        End Select
'        If Result = 0 Then GoTo ErrCrystal
'    Next
'    '--------------------------------------------------------------------------------------------------------------------------------------------
'
'    'Seteo la Query del reporte-----------------------------------------------------------------
'    Cons = "SELECT Documento.DocCodigo , Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal, Documento.DocIVA, Documento.DocVendedor" _
'            & " From " & paBD & ".dbo.Documento Documento " _
'            & " Where DocCodigo = " & aRecibo
'    If crSeteoSqlQuery(jobnum%, Cons) = 0 Then GoTo ErrCrystal
'
'    'Subreporte srContado.rpt  y srContado.rpt - 01-----------------------------------------------------------------------------
'    JobSRep1 = crAbroSubreporte(jobnum, "srRecibo.rpt")
'    If JobSRep1 = 0 Then GoTo ErrCrystal
'
'    Cons = "SELECT  DocumentoPago.DPaDocASaldar, DocumentoPago.DPaDocQSalda, DocumentoPago.DPaCuota, DocumentoPago.DPaDe, " _
'                       & " DocumentoPago.DPaAmortizacion, DocumentoPago.DPaMora, Documento.DocSerie, Documento.DocNumero, " _
'                       & " CreditoCuota.CCuValor, CreditoCuota.CCuVencimiento, Credito.CreProximoVto " _
'            & " From { oj ((" & paBD & ".dbo.DocumentoPago DocumentoPago " _
'                                & " INNER JOIN " & paBD & ".dbo.Documento Documento ON DocumentoPago.DPaDocASaldar = Documento.DocCodigo)" _
'                                & " LEFT OUTER JOIN " & paBD & ".dbo.Credito Credito ON Documento.DocCodigo = Credito.CreFactura)" _
'                                & " LEFT OUTER JOIN " & paBD & ".dbo.CreditoCuota CreditoCuota ON Credito.CreCodigo = CreditoCuota.CCuCredito And DocumentoPago.DPaCuota = CreditoCuota.CCuNumero}" _
'            & " Where DocumentoPago.DPaDocQSalda = " & aRecibo
'
'
'    If crSeteoSqlQuery(JobSRep1, Cons) = 0 Then GoTo ErrCrystal
'
'    JobSRep2 = crAbroSubreporte(jobnum, "srRecibo.rpt - 01")
'    If JobSRep2 = 0 Then GoTo ErrCrystal
'    If crSeteoSqlQuery(JobSRep2, Cons) = 0 Then GoTo ErrCrystal
'    '-------------------------------------------------------------------------------------------------------------------------------------
'
'    'If crMandoAPantalla(jobnum, "Recibo de Pago") = 0 Then GoTo ErrCrystal
'    If crMandoAImpresora(jobnum, 1) = 0 Then GoTo ErrCrystal
'    If Not crInicioImpresion(jobnum, True, False) Then GoTo ErrCrystal
'
'    If Not crCierroSubReporte(JobSRep1) Then GoTo ErrCrystal
'    If Not crCierroSubReporte(JobSRep2) Then GoTo ErrCrystal
'
'    'crEsperoCierreReportePantalla
'
'    crCierroTrabajo jobnum
'
'    Screen.MousePointer = 0
'    Exit Sub
'
'ErrCrystal:
'    Screen.MousePointer = 0
'    clsGeneral.OcurrioError crMsgErr, Err.Description
'    On Error Resume Next
'    Screen.MousePointer = 11
'    crCierroSubReporte JobSRep1
'    crCierroSubReporte JobSRep2
'    crCierroTrabajo jobnum
'    Screen.MousePointer = 0
'End Sub

'Private Function ImprimoMoras(txtDatos As String, Optional TOPrinter As Boolean = True, _
'            Optional Cliente As String = "", Optional Cedula As String = "", _
'            Optional RUC As String = "", Optional ByVal Usuario As Long = 0, Optional EsRedPagos As Boolean = False)
'
'Dim job_Moras As Integer, job_R As Integer, job_QFrms As Integer, job_FrmName As String
'
'    On Error GoTo ErrCrystal
'
'    'Valores para las fx del Reporte -----------------------------------------------------------------------------
'    If Cliente = "" Then Cliente = Trim(lNombre.Caption)
'    If Cedula = "" Then
'        If Val(lCiRuc.Tag) = TipoCliente.Persona Then Cedula = Trim(tCliente.Text)
'    End If
'    If RUC = "" Then
'        If Val(lCiRuc.Tag) = TipoCliente.Empresa Then RUC = Trim(tCliente.Text) Else RUC = Trim(lRuc.Tag)
'    End If
'    If Usuario = 0 Then Usuario = paCodigoDeUsuario
'    '-------------------------------------------------------------------------------------------------------------
'
'    job_Moras = crAbroReporte(prmPathListados & "Aporte.RPT")
'    If job_Moras = 0 Then GoTo ErrCrystal
'
'    job_QFrms = crObtengoCantidadFormulasEnReporte(job_Moras)
'
'    Dim oPrintAux As New clsPrintRedPagos
'    oPrintAux.Impresora = paIReciboN
'    oPrintAux.Bandeja = paIReciboB
'
'    If job_QFrms = -1 Then GoTo ErrCrystal
'
'    If Not crSeteoImpresora(job_Moras, Printer, oPrintAux.Bandeja, mOrientation:=2, PaperSize:=13) Then GoTo ErrCrystal    'PAPEL A5
'
'    Dim arrWork() As String, arrValor() As String, idx As Integer
'    arrWork = Split(txtDatos, ";")
'
'    For idx = LBound(arrWork) To UBound(arrWork)
'        arrValor = Split(arrWork(idx), ":")     '0-idDoc    1-Moneda
'
'        Screen.MousePointer = 11
'
'        'Cargo Propiedades para el reporte Contado --------------------------------
'        For i = 0 To job_QFrms - 1
'            job_FrmName = crObtengoNombreFormula(job_Moras, i)
'
'            Select Case LCase(job_FrmName)
'                Case "": GoTo ErrCrystal
'                Case "nombredocumento": job_R = crSeteoFormula(job_Moras, job_FrmName, "'" & paDNDebito & "'")
'
'                Case "cliente": job_R = crSeteoFormula(job_Moras%, job_FrmName, "'" & Cliente & "'")
'                Case "cedula": job_R = crSeteoFormula(job_Moras%, job_FrmName, "'" & Cedula & "'")
'                Case "ruc": job_R = crSeteoFormula(job_Moras%, job_FrmName, "'" & RUC & "'")
'
'                Case "signomoneda": job_R = crSeteoFormula(job_Moras%, job_FrmName, "'" & BuscoSignoMoneda(Val(arrValor(1))) & "'")
'                Case "nombremoneda": job_R = crSeteoFormula(job_Moras%, job_FrmName, "'(" & BuscoNombreMoneda(Val(arrValor(1))) & ")'")
'
'                Case "usuario": job_R = crSeteoFormula(job_Moras%, job_FrmName, "'" & BuscoDigitoUsuario(Usuario) & "'")
'
'                Case "cuenta": job_R = crSeteoFormula(job_Moras%, job_FrmName, "'" & UCase("Concepto: Intereses por Mora") & "'")
'                Case "articulo": job_R = crSeteoFormula(job_Moras%, job_FrmName, "''")
'
'                Case Else: job_R = 1
'            End Select
'            If job_R = 0 Then GoTo ErrCrystal
'        Next
'        '--------------------------------------------------------------------------------------------------------------------------------------------
'
'        'Seteo la Query del reporte-----------------------------------------------------------------
'        Cons = "SELECT Documento.DocCodigo , Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal, Documento.DocIVA, Documento.DocVendedor" _
'                & " From " & paBD & ".dbo.Documento Documento " _
'                & " Where DocCodigo = " & arrValor(0)
'        If crSeteoSqlQuery(job_Moras%, Cons) = 0 Then GoTo ErrCrystal
'
'        If Not TOPrinter Then If crMandoAPantalla(job_Moras, "Nota de Debito") = 0 Then GoTo ErrCrystal
'        If TOPrinter Then If crMandoAImpresora(job_Moras, 1) = 0 Then GoTo ErrCrystal
'
'        If Not crInicioImpresion(job_Moras, True, False) Then GoTo ErrCrystal
'
'        If Not TOPrinter Then crEsperoCierreReportePantalla
'    Next
'
'    crCierroTrabajo job_Moras
'    Screen.MousePointer = 0
'    Exit Function
'
'ErrCrystal:
'    Screen.MousePointer = 0
'    clsGeneral.OcurrioError crMsgErr
'    On Error Resume Next
'    Screen.MousePointer = 11
'    crCierroTrabajo job_Moras
'    Screen.MousePointer = 0
'End Function

'Private Sub ImprimoReciboSeniaONotaDebito(ByVal aDocumento As Long, ByVal Articulo As String, ByVal NombreDocumento As String)
'
'Dim Result As Integer
'Dim NombreFormula As String
'
'    On Error GoTo ErrCrystal
'    If Not InicializoCrystalEngine("Aporte.RPT") Then Screen.MousePointer = 0: Exit Sub
'
'    Screen.MousePointer = 11
'
'    For i = 0 To CantForm - 1
'        NombreFormula = crObtengoNombreFormula(jobnum, i)
'
'        Select Case LCase(NombreFormula)
'            Case "": GoTo ErrCrystal
'            'paDRecibo
'            Case "nombredocumento": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & NombreDocumento & "'")
'
'            Case "cliente": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lTitular.Caption) & "'")
'            Case "cedula":
'                If Val(lCiRuc.Tag) = 1 Then
'                    Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lCiRuc.Caption) & "'")
'                Else
'                    Result = crSeteoFormula(jobnum%, NombreFormula, "''")
'                End If
'
'            Case "ruc":
'                If Val(lCiRuc.Tag) = TipoCliente.Empresa Then
'                    Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lCiRuc.Caption) & "'")
'                Else
'                    Result = crSeteoFormula(jobnum%, NombreFormula, "''")
'                End If
'
'            Case "signomoneda": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lMoneda.Caption) & "'")
'            Case "nombremoneda": Result = crSeteoFormula(jobnum%, NombreFormula, "'(" & BuscoNombreMoneda(lMoneda.Tag) & ")'")
'
'            Case "usuario": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & BuscoDigitoUsuario(paCodigoDeUsuario) & "'")
'
'            Case "cuenta": Result = crSeteoFormula(jobnum%, NombreFormula, "''")
'
'            Case "articulo":
'                'aTexto = "Devolución de Seña"
'                Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Articulo & "'")
'
'            Case Else: Result = 1
'        End Select
'        If Result = 0 Then GoTo ErrCrystal
'    Next
'    '--------------------------------------------------------------------------------------------------------------------------------------------
'
'    'Seteo la Query del reporte-----------------------------------------------------------------
'    Cons = "SELECT Documento.DocCodigo , Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal, Documento.DocIVA, Documento.DocVendedor" _
'            & " From " & paBD & ".dbo.Documento Documento " _
'            & " Where DocCodigo = " & aDocumento
'    If crSeteoSqlQuery(jobnum%, Cons) = 0 Then GoTo ErrCrystal
'    '-------------------------------------------------------------------------------------------------------------------------------------
'
'    'If crMandoAPantalla(JobNum, "Recibo de Pago") = 0 Then GoTo ErrCrystal
'    If crMandoAImpresora(jobnum, 1) = 0 Then GoTo ErrCrystal
'    If Not crInicioImpresion(jobnum, True, False) Then GoTo ErrCrystal
'
'    'crEsperoCierreReportePantalla
'    crCierroTrabajo jobnum
'    Screen.MousePointer = 0
'    Exit Sub
'
'ErrCrystal:
'    clsGeneral.OcurrioError crMsgErr, Err.Description
'    Screen.MousePointer = 0
'End Sub

Private Function InicializoCrystalEngine(Reporte As String) As Boolean
    
    'Inicializa el Engine del Crystal y setea la impresora para el JOB
    On Error GoTo ErrCrystal
    InicializoCrystalEngine = False
    
    'Inicializo el Reporte y SubReportes
    jobnum = crAbroReporte(prmPathListados & Reporte)
    If jobnum = 0 Then GoTo ErrCrystal
    
    'Configuro la Impresora
    If Trim(Printer.DeviceName) <> Trim(paIReciboN) Then SeteoImpresoraPorDefecto paIReciboN
    If Not crSeteoImpresora(jobnum, Printer, paIReciboB) Then GoTo ErrCrystal

    'Obtengo la cantidad de formulas que tiene el reporte.
    CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
    If CantForm = -1 Then GoTo ErrCrystal
    InicializoCrystalEngine = True
    Exit Function

ErrCrystal:
    clsGeneral.OcurrioError Trim(crMsgErr) & " No se podrán imprimir Recibos de Pago.", Err.Description
    Screen.MousePointer = 0
End Function

Private Function VerificoCuentaDocumento(IdDocumento As Long) As Boolean

    On Error GoTo errVerifico
    Dim rsVer
    VerificoCuentaDocumento = True
    Cons = "Select * from CuentaDocumento Where CDoIdDocumento = " & IdDocumento
    Set rsVer = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If rsVer.EOF Then VerificoCuentaDocumento = False
    rsVer.Close
    Exit Function

errVerifico:
    clsGeneral.OcurrioError "Ocurrió un error al verificar las tablas de colectivos.", Err.Description
End Function

Function BuscoNombreMoneda(Codigo As Long) As String

    On Error GoTo ErrBU
    Dim Rs As rdoResultset
    BuscoNombreMoneda = ""

    Cons = "Select * from Moneda WHERE MonCodigo = " & Codigo
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not Rs.EOF Then BuscoNombreMoneda = Trim(Rs!MonNombre)
    Rs.Close
    Exit Function
    
ErrBU:
End Function

Function BuscoDigitoUsuario(Codigo As Long) As String
On Error GoTo ErrBU
Dim Rs As rdoResultset

    BuscoDigitoUsuario = ""
    Cons = "Select * from Usuario Where UsuCodigo = " & Codigo
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not Rs.EOF Then BuscoDigitoUsuario = Trim(Rs!UsuDigito)
    Rs.Close
    Exit Function
    
ErrBU:
End Function

Private Function ReciboConMoraANotaMultiple(ByVal xIDRecibo As Long) As Boolean
On Error GoTo errValidar
    ReciboConMoraANotaMultiple = True
    Dim Rs As rdoResultset
    'Controlo si la Nota ya fue emitida --> Estado = 1
    Cons = "Select TOP 1 * from NotasSinFacturar Where NSFEstado = 0 AND NSFIDRecibo = " & xIDRecibo
    Set Rs = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    ReciboConMoraANotaMultiple = Not Rs.EOF
    Rs.Close
    Exit Function
errValidar:
    clsGeneral.OcurrioError "Error al validar si el recibo está asignado a una nota de débito múltiple.", Err.Description
    Screen.MousePointer = 0
End Function


Private Function ValidoReciboEnNotaMultiple(xIDRecibo As Long) As Boolean
On Error GoTo errValidar
    ValidoReciboEnNotaMultiple = True
    
    'Controlo si la Nota ya fue emitida --> Estado = 1
    Cons = "Select TOP 1 * from NotasSinFacturar Where NSFEstado = 1 AND NSFIDRecibo = " & xIDRecibo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    ValidoReciboEnNotaMultiple = Not RsAux.EOF
    RsAux.Close
    
    Exit Function
errValidar:
    clsGeneral.OcurrioError "Error al validar si el recibo está asignado a una nota de débito múltiple.", Err.Description
    Screen.MousePointer = 0
End Function

Private Function ValidoCajaCerrada() As Boolean
On Error GoTo errCierre
Dim bEstaCerrada As Boolean

    ValidoCajaCerrada = True
    
    'Valido si hay un cierre de caja para la disponibilidad
    Cons = "Select  * from MovimientoDisponibilidad, MovimientoDisponibilidadRenglon" _
           & " Where MDiID = MDRIdMovimiento " _
           & " And MDiFecha = '" & Format(gFechaServidor, "mm/dd/yyyy") & "'" _
           & " And MDiHora = '23:59:59'" _
           & " And MDiTipo = " & paMCIngresosOperativos _
           & " And MDRIdDisponibilidad = " & paDisponibilidad
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    bEstaCerrada = Not RsAux.EOF
    RsAux.Close
    
    If bEstaCerrada Then
        If MsgBox("La caja para el día " & Format(gFechaServidor, "d Mmmm") & " está cerrada." & vbCrLf & _
                       "Si continúa con la anulación recuerde volver a realizar el cierre." & vbCrLf & vbCrLf & _
                       "Está seguro de continuar ? ", vbQuestion + vbYesNo + vbDefaultButton2, "La Caja está Cerrada") = vbNo Then
            ValidoCajaCerrada = False
        End If
    End If
    
errCierre:
End Function

Private Function ValidoDisponibilidad() As Long
    
    Dim mRet As Long
    ValidoDisponibilidad = 0
    mRet = dis_DisponibilidadPara(paCodigoDeSucursal, CLng(lMoneda.Tag))
    
    If mRet = 0 Then
        MsgBox "Ud. no puede anular éste documento." & vbCrLf & _
                    "Su sucursal no tiene asignada una disponibilidad para hacer los movimientos en la moneda del documento." & vbCr & vbCr & _
                    "Consulte con el administrador del sistema.", vbExclamation, "Falta Disponibilidad "
    End If
    
    ValidoDisponibilidad = mRet
    
End Function
Private Function EmitirCFE(ByVal Documento As clsDocumentoCGSA, ByVal CAE As clsCAEDocumento, ByVal codSucursalDGI As String) As String
On Error GoTo errEC
    
    If (EmpresaEmisora Is Nothing) Then
        Set EmpresaEmisora = New clsClienteCFE
        EmpresaEmisora.CargoInformacionCliente cBase, 1, False
    End If
    
    If prmURLFirmaEFactura = "" Then
        CargoValoresIVA
        CargarParametroEFactura
    End If
    
    With New clsCGSAEFactura
        .URLAFirmar = prmURLFirmaEFactura
        .TasaBasica = TasaBasica
        .TasaMinima = TasaMinima
        .ImporteConInfoDeCliente = prmImporteConInfoCliente
        Set .Connect = cBase
        If Not .GenerarEComprobante(CAE, Documento, EmpresaEmisora, codSucursalDGI) Then
            EmitirCFE = .XMLRespuesta
        End If
    End With
    Exit Function
    
errEC:

End Function
Private Function FirmoCFE(ByVal Documento As Long) As String
On Error GoTo errEC
    With New clsCGSAEFactura
        .URLAFirmar = prmURLFirmaEFactura
        .TasaBasica = TasaBasica
        .TasaMinima = TasaMinima
        .ImporteConInfoDeCliente = prmImporteConInfoCliente
        Set .Connect = cBase
        Dim sFirma As String
        FirmoCFE = IIf(LCase(.FirmarUnDocumento(Documento)) = "false", .XMLRespuesta, "")
    End With
    Exit Function
errEC:
    FirmoCFE = "Error en firma: " & Err.Description
End Function
Private Sub CargoArticuloInteresesPorMora()
Dim rsArtMora As rdoResultset
        
    Set ProdInteresMora = New clsProducto
    Set rsArtMora = cBase.OpenResultset("SELECT ArtCodigo, ArtID, ArtTipo, ArtNombre, IvaCodigo, IvaDescripcion, IvaPorcentaje  FROM Articulo " _
                & "INNER JOIN ArticuloFacturacion ON ArtID = AFaArticulo " _
                & "INNER JOIN TipoIVA ON AFaIva = IvaCodigo " _
                & "WHERE ArtID = " & 5547, rdOpenDynamic, rdConcurValues)
                
    With ProdInteresMora
        .ID = 5547
        .Nombre = Trim(rsArtMora("ArtNombre"))
        .TipoArticulo = rsArtMora("ArtTipo")
        .TipoIVA.Porcentaje = rsArtMora("IvaPorcentaje")
    End With
    rsArtMora.Close
        
End Sub

Private Function RetornoDocsRelacionados() As Collection
Dim CodAnt As Long
    CodAnt = 0
    Set RetornoDocsRelacionados = New Collection
    For Each itmX In lRecibo.ListItems
        If Val(Mid(itmX.Key, 2)) = 2 Then
            If CodAnt <> CLng(itmX.Tag) Then
                Dim oDocRel As New clsDocumentoAsociado
                RetornoDocsRelacionados.Add oDocRel
                With oDocRel
                    .ID = CLng(itmX.Tag)
                    .Fecha = itmX.SubItems(5)
                    .Numero = Mid(itmX.Text, 2)
                    .Serie = Mid(itmX.Text, 1, 1)
                    .Tipo = TD_Credito
                    .TipoEFactura = IIf(Val(lTitular.Tag) = 1, CFE_eFactura, CFE_eTicket)
                End With
                CodAnt = CLng(itmX.Tag)
            End If
            oDocRel.Devuelve = oDocRel.Devuelve + CCur(itmX.SubItems(3))
        End If
    Next
                
End Function

