VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{190700F0-8894-461B-B9F5-5E731283F4E1}#1.1#0"; "orHiperlink.ocx"
Begin VB.Form frmRetiroConCompra 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Retiro en domicilio"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8520
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRetiroConCompra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCompra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   0
      ScaleHeight     =   3705
      ScaleWidth      =   8505
      TabIndex        =   14
      Top             =   3960
      Width           =   8535
      Begin VB.TextBox txtMemo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   80
         TabIndex        =   23
         Top             =   2160
         Width           =   7095
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   2520
         Width           =   735
      End
      Begin VB.CheckBox chkIncluirAlPendiente 
         BackColor       =   &H00FFFFFF&
         Caption         =   "¿El camionero tiene que cobrar en domicilio la nueva compra?"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   6735
      End
      Begin VB.TextBox txtEnvioCompra 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   840
         MaxLength       =   10
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin prjHiperLink.orHiperLink hliClienteEnvio 
         Height          =   315
         Left            =   840
         Top             =   960
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   556
         BackColor       =   16777215
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
      Begin prjHiperLink.orHiperLink hliDocEnvio 
         Height          =   315
         Left            =   2160
         Top             =   600
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   556
         BackColor       =   16777215
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
         MouseIcon       =   "frmRetiroConCompra.frx":2B05
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Comentario:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "&Usuario"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Envío de nueva compra"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   8295
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Envío:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.PictureBox picPasoNota 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4140
      Left            =   0
      ScaleHeight     =   4110
      ScaleWidth      =   8490
      TabIndex        =   7
      Top             =   600
      Width           =   8520
      Begin VB.CheckBox chkNotaPendiente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "¿El dinero se le devuelve en el domicilio al cliente?"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   7215
      End
      Begin VB.TextBox txtDocumento 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   8
         Top             =   600
         Width           =   1815
      End
      Begin prjHiperLink.orHiperLink hliCliente 
         Height          =   315
         Left            =   960
         Top             =   960
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   556
         BackColor       =   16777215
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
         Height          =   315
         Left            =   4080
         Top             =   600
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         BackColor       =   16777215
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
         MouseIcon       =   "frmRetiroConCompra.frx":2E1F
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
         Height          =   1695
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   8175
         _cx             =   14420
         _cy             =   2990
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
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Artículos devueltos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Nota con ficha de devolución"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   8295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Artículos A RETIRAR"
         ForeColor       =   &H00000000&
         Height          =   15
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblDocumento 
         BackStyle       =   0  'Transparent
         Caption         =   "&Documento, C.I./R.U.C."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1020
         Width           =   1095
      End
   End
   Begin VB.PictureBox picTitulo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   8490
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   8520
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   120
         Picture         =   "frmRetiroConCompra.frx":3139
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   6
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Retiro con compra en domicilio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Top             =   120
         Width           =   5175
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
      ScaleWidth      =   8490
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7785
      Width           =   8520
      Begin VB.CommandButton butAtras 
         Caption         =   "&Atrás"
         Height          =   375
         Left            =   4680
         TabIndex        =   26
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton butCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   7320
         TabIndex        =   3
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton butAceptar 
         Caption         =   "&Siguiente"
         Height          =   375
         Left            =   6000
         TabIndex        =   2
         Top             =   120
         Width           =   1215
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
         Width           =   4500
      End
   End
End
Attribute VB_Name = "frmRetiroConCompra"
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

Private iPasoWiz As Integer

Public Sub SetearControlesWizard()
On Error Resume Next

    picCompra.Visible = (iPasoWiz = 1)
    picPasoNota.Visible = (iPasoWiz = 0)
    butAtras.Enabled = (iPasoWiz > 0)
    
    Select Case iPasoWiz
        Case 0:
            butAceptar.Caption = "Siguiente"
            butAceptar.Enabled = (Val(hliDocumento.Tag) > 0)
            txtDocumento.SetFocus
        Case 1:
            butAceptar.Caption = "Finalizar"
            butAceptar.Enabled = (Val(txtEnvioCompra.Tag) > 0)
    End Select
    
End Sub

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

Private Sub loc_InsertDocumentoPendiente(ByVal lDoc As Long, ByVal iTipo As Integer, ByVal lIDTipo As Long, ByVal cImporte As Currency, ByVal iMon As Integer)
Dim m_Disponibilidad As Long
    m_Disponibilidad = modMaeDisponibilidad.dis_DisponibilidadPara(CLng(paCodigoDeSucursal), CLng(iMon))
    Cons = "Insert Into DocumentoPendiente (DPeDocumento, DPeTipo, DPeIDTipo, DPeImporte, DPeMoneda, DPeDisponibilidad) Values (" & _
                    lDoc & ", " & iTipo & ", " & lIDTipo & ", " & Format(cImporte, "###0.00") & ", " & iMon & ", " & m_Disponibilidad & ")"
    cBase.Execute Cons
End Sub


Private Function GrabarRemitosEnBD(ByVal cliente As clsClienteCFE) As Boolean
   
   GrabarRemitosEnBD = False
   Dim vInf() As String
   vInf = Split(hliDocEnvio.Tag, "|")
   
    If (chkIncluirAlPendiente.value) And Val(chkIncluirAlPendiente.Tag) = 2 Then
         Cons = "SELECT COUNT(DocCodigo), MAX(DocFecha) FROM Documento INNER JOIN DocumentoPago ON DPADocQSalda = DocCodigo " & _
                     " WHERE DPaDocASaldar = " & vInf(1) & " AND DocTipo = 5"
         Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
         If Not rsAux.EOF Then
            If rsAux(0) <> 1 Then
                 MsgBox "Sólo se admite incluir créditos que posean un único recibo impreso.", vbExclamation, "ATENCIÓN"
                 Exit Function
            ElseIf Format(rsAux(1), "dd/MM/yyyy") <> Format(Now, "dd/MM/yyyy") Then
                MsgBox "Sólo se admite incluir en el pendiente documentos impresos en el día.", vbExclamation, "ATENCIÓN"
                 Exit Function
            End If
        End If
        rsAux.Close
        
    End If
    
    If (chkNotaPendiente.value) Then
        Cons = "SELECT DocFecha FROM Documento WHERE DocCodigo = " & Val(hliDocumento.Tag)
        Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Format(rsAux(0), "dd/MM/yyyy") <> Format(Now, "dd/MM/yyyy") Then
            rsAux.Close
            MsgBox "La nota no es del día .", vbExclamation, "ATENCIÓN"
            Exit Function
        End If
        rsAux.Close
    End If
   
    On Error GoTo errBT
   cBase.BeginTrans
   On Error GoTo ErrResumo
    
    Dim caeG As New clsCAEGenerador
    Dim CAE As clsCAEDocumento
    Set CAE = caeG.ObtenerNumeroCAEDocumento(cBase, CGSA_TiposCFE.CFE_eRemito, paCodigoDGI)
    
    Dim docRemitoRet As clsDocumentoCGSA
    Set docRemitoRet = New clsDocumentoCGSA
        
    Dim oRenDoc As clsDocumentoRenglon
    Dim I As Integer
    For I = 1 To lstArticulos.Rows - 1
        Set oRenDoc = New clsDocumentoRenglon
        
        oRenDoc.Cantidad = Val(lstArticulos.Cell(flexcpText, I, 0))
        oRenDoc.CantidadARetirar = oRenDoc.Cantidad
        oRenDoc.EstadoMercaderia = paEstadoArticuloEntrega
        oRenDoc.IVA = 0
        oRenDoc.Precio = 0
        oRenDoc.Articulo.ID = lstArticulos.Cell(flexcpData, I, 0)
        
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
        .Comentario = "Retiro en domicilio en envío: " & Val(txtEnvioCompra.Tag) & ". " & txtMemo.Text
        .Vendedor = Val(txtUser.Tag)
    End With
    Set docRemitoRet.Conexion = cBase
    docRemitoRet.Codigo = docRemitoRet.InsertoDocumentoBD(0)
        
    For I = 1 To lstArticulos.Rows - 1
        Cons = "SELECT * FROM Devolucion WHERE DevID = " & lstArticulos.Cell(flexcpText, I, 1)
        Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        rsAux.Edit
        rsAux("DevRemito") = docRemitoRet.Codigo
        rsAux.Update
        rsAux.Close
    Next
    
    Dim rsEnv As rdoResultset
    Dim sQy As String
    sQy = "SELECT * FROM Envio WHERE EnvCodigo = " & Val(txtEnvioCompra.Text) & " AND EnvEstado = 0"
    Set rsEnv = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    cBase.Execute "INSERT INTO EnviosRemitos (EReEnvio, EReRemito, EReFactura, EReNota) VALUES (" & rsEnv("EnvCodigo") & ", " & docRemitoRet.Codigo & ", " & rsEnv("EnvDocumento") & ", " & Val(hliDocumento.Tag) & ")"
    
    If (chkNotaPendiente.value) Then
        loc_InsertDocumentoPendiente Val(hliDocumento.Tag), 1, rsEnv("EnvCodigo"), CCur(chkNotaPendiente.Tag) * -1, rsEnv("EnvMoneda")
    End If
    
    If (chkIncluirAlPendiente.value) Then
        
        If Val(chkIncluirAlPendiente.Tag) = 1 Then
            Cons = "SELECT DocCodigo, DocTotal, DocMoneda FROM Documento WHERE DocCodigo = " & vInf(1)
        Else
            Cons = "SELECT DocCodigo, DocTotal, DocMoneda FROM Documento INNER JOIN DocumentoPago ON DPADocQSalda = DocCodigo " & _
                " WHERE DPaDocASaldar = " & vInf(1) & " AND DocTipo = 5 ORDER BY DocCodigo DESC"
        End If
        Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        'Si es crédito pongo el recibo en pendiente.
        loc_InsertDocumentoPendiente rsAux("DocCodigo"), 1, rsEnv("EnvCodigo"), rsAux("DocTotal"), rsAux("DocMoneda")
        rsAux.Close
        
    End If
    
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
    Cons = "SELECT ArtID, ArtCodigo, ArtNombre, DevCantidad Cantidad, DevID Fichas " & _
        " FROM Documento " & _
        " INNER JOIN Renglon ON RenDocumento = DocCodigo " & _
        " INNER JOIN Articulo ON ArtID = RenArticulo AND ArtTipo <> 151" & _
        " INNER JOIN Devolucion ON DocCodigo = DevNota AND DevNota IS NOT NULL AND DevLocal IS NULL AND DevAnulada IS NULL AND DevArticulo = RenArticulo AND DevRemito IS NULL " & _
        " WHERE DocCodigo = " & hliDocumento.Tag & " ORDER BY ArtNombre "
    Set rsArt = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsArt.EOF
        With lstArticulos
            .AddItem rsArt("Cantidad")
            .Cell(flexcpText, .Rows - 1, 1) = rsArt("Fichas")
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsArt("ArtCodigo"), "#,##0,000") & " " & Trim(rsArt("ArtNombre"))
            .Cell(flexcpData, .Rows - 1, 0) = CStr(rsArt(rsArt("ArtID")))
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
            Case txtEnvioCompra.Name
                lblAyuda.Caption = "Ingrese el código del envío de compra para asociar a la nota."
            Case chkIncluirAlPendiente.Name
                lblAyuda.Caption = "Si no es venta telefónica se pondrá el documento en pendiente."
        End Select
    Else
        lblAyuda.Caption = msg
    End If
End Sub

Private Sub LimpiarControlesArticulos()
    
    txtMemo.Text = ""
    txtUser.Text = "": txtUser.Tag = 0
    lstArticulos.Rows = 1
    butAceptar.Enabled = False
    
End Sub

Private Sub LimpiarControlesDocumento()
    hliCliente.Caption = ""
    hliDocumento.Caption = ""
    hliCliente.Tag = ""
    hliDocumento.Tag = ""
    chkNotaPendiente.value = 0
    chkNotaPendiente.Tag = ""
    LimpiarControlesEnvio
End Sub

Private Sub LimpiarControlesEnvio()
    hliClienteEnvio.Caption = ""
    hliDocEnvio.Caption = ""
    hliClienteEnvio.Tag = ""
    hliDocEnvio.Tag = ""
    txtEnvioCompra.Tag = ""
    chkIncluirAlPendiente.value = 0
End Sub


Private Sub butAceptar_Click()
    
    If iPasoWiz = 0 Then
    
        iPasoWiz = 1
        SetearControlesWizard
        Exit Sub
        
    End If
    
    On Error GoTo errValidar
    
    If Val(txtUser.Tag) = 0 Then
        MsgBox "Ingrese su dígito de usuario.", vbExclamation, "Validación"
        txtUser.SetFocus
        Exit Sub
    End If
    
    If lstArticulos.Rows - 1 = 0 Then
        MsgBox "No hay artículos para retirar.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    Dim oInfoCAE As New clsCAEGenerador
    If Not oInfoCAE.SucursalTieneCae(cBase, CGSA_TiposCFE.CFE_eRemito, paCodigoDGI) Then
        MsgBox "No hay un CAE disponible para emitir el eRemito de retiro, por favor comuníquese con administración." & vbCrLf & vbCrLf & "No podrá grabar.", vbCritical, "eFactura"
        Screen.MousePointer = 0
        Exit Sub
    End If
       
    'Si es un remito recepción tengo que obligar a ingresar todos los artículos.
    If MsgBox("¿Confirma almacenar los datos ingresados?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
        
        If (EmpresaEmisora Is Nothing) Then
            Set EmpresaEmisora = New clsClienteCFE
            EmpresaEmisora.CargoClienteCarlosGutierrez paCodigoDeSucursal
        End If

        If Not (chkNotaPendiente.value) Then
            If (MsgBox("NO MARCO LA NOTA AL PENDIENTE!!!!" & vbCrLf & vbCrLf & "La nota es una salida de caja, el dinero se le entregó al cliente?" & vbCrLf & vbCrLf & "¿Confirma continuar?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "POSIBLE ERROR") <> vbYes) Then
                Exit Sub
            End If
        End If
        FechaDelServidor
        If GrabarRemitosEnBD(EmpresaEmisora) Then
            Call butCancelar_Click
        End If
    End If
    Exit Sub
    
errValidar:
    clsGeneral.OcurrioError "Error al validar para grabar", Err.Description
    Exit Sub

End Sub

Private Sub butAtras_Click()
On Error Resume Next
    iPasoWiz = iPasoWiz - 1
    SetearControlesWizard
End Sub

Private Sub butCancelar_Click()
    
    'cancelo el ingreso.
    LimpiarControlesDocumento
    LimpiarControlesArticulos
    
    iPasoWiz = 0
    butAceptar.Enabled = False
    SetearControlesWizard
    
    txtDocumento.Text = ""
    txtDocumento.SetFocus
    
End Sub

Private Sub chkIncluirAlPendiente_GotFocus()
    MostrarAyuda
End Sub

Private Sub chkIncluirAlPendiente_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then txtMemo.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
Dim sPaso As String

    sPaso = "Controles"
    LimpiarControlesArticulos
    LimpiarControlesDocumento

    Me.Height = 5805
    picCompra.Move picPasoNota.Left, picPasoNota.Top, picPasoNota.Width, picPasoNota.Height
    
    With lstArticulos
        .Rows = 1: .Cols = 1
        .RowHeight(0) = 315
        .RowHeightMin = 285
        .FixedCols = 0
        .FormatString = ">Devuelve|>Ficha|<Artículo"
        .ColWidth(0) = 1000
        .ColWidth(1) = 1000
        .ColWidth(2) = 2000
        .ExtendLastCol = True
    End With
        SetearControlesWizard
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
    
    Exit Sub
    'If lstArticulos.Row < 1 Then Exit Sub
    
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

Private Sub txtDocumento_Change()
    If Val(hliDocumento.Tag) > 0 Or Val(hliCliente.Tag) > 0 Then
        LimpiarControlesDocumento
        LimpiarControlesArticulos
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
        
        If InStr(1, txtDocumento.Text, "d", vbTextCompare) > 0 Then
            Cons = Replace("SELECT DocCodigo, CliCodigo, DocFModificacion, dbo.NombreTipoDocumento(100+DocTipo) + ' ' + rTrim(DocSerie)+'-'+Convert(VarChar(10), DocNumero) Documento, " & _
                "RTrim(IsNull(CEmFantasia, rTrim(CPeApellido1) + ', ' + RTrim(CPeNombre1))) Cliente, " & _
                "ISNULL(CliCiRUC, '') [C.I./R.U.C.]  " & _
                "FROM Documento " & _
                "INNER JOIN CLiente ON DocCliente = CliCodigo " & _
                "LEFT OUTER JOIN CPersona ON CPeCliente = CliCodigo LEFT OUTER JOIN CEmpresa ON CEmCliente = CliCodigo " & _
                "WHERE DocTipo = SUBSTRING(@RUC, 1, CHARINDEX('d', @RUC)-1) AND DocTipo IN (3,4,10) AND DocCodigo = CONVERT(int, SUBSTRING(@RUC, CHARINDEX('d', @RUC)+1, 10)) " & _
                "AND DocAnulado = 0 ORDER BY DocFecha DESC", "@RUC", txtDocumento.Text)
        ElseIf IsNumeric(txtDocumento.Text) Then
            Cons = Replace("SELECT 0 DocCodigo, IsNull(CliCodigo, 0) CliCodigo, GetDate() DocFModificacion, '' Documento, " & _
                "RTrim(IsNull(CEmFantasia, rTrim(CPeApellido1) + ', ' + RTrim(CPeNombre1))) Cliente " & _
                ", rTrim(CliCIRuc) [C.I./R.U.C.] " & _
                "FROM ((Cliente LEFT OUTER JOIN CEmpresa ON CliCodigo = CEmCliente) LEFT OUTER JOIN CPersona ON CliCodigo = CPeCliente) Where CliCiRuc LIKE @RUC ", "@RUC", txtDocumento.Text)
        Else
            Cons = Replace("SELECT DocCodigo, CliCodigo, DocFModificacion, dbo.NombreTipoDocumento(100+DocTipo) + ' ' + rTrim(DocSerie)+'-'+Convert(VarChar(10), DocNumero) Documento, " & _
                "RTrim(IsNull(CEmFantasia, rTrim(CPeApellido1) + ', ' + RTrim(CPeNombre1))) Cliente, " & _
                "ISNULL(CliCiRUC, '') [C.I./R.U.C.]  " & _
                "FROM Documento " & _
                "INNER JOIN CLiente ON DocCliente = CliCodigo " & _
                "LEFT OUTER JOIN CPersona ON CPeCliente = CliCodigo LEFT OUTER JOIN CEmpresa ON CEmCliente = CliCodigo " & _
                "WHERE DocTipo IN (3,4,10) AND DocSerie = SUBSTRING(@RUC, 1, 1) AND DocNumero = CONVERT(int, SUBSTRING(@RUC, 2, 10)) " & _
                "AND DocAnulado = 0 ORDER BY DocFecha DESC", "@RUC", "'" & txtDocumento.Text & "'")
        End If
        
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
                
            'es seleccionado por lista de ayuda.
            Cons = "SELECT DocTotal, DocCodigo, CliCodigo, DocFModificacion, dbo.NombreTipoDocumento(100+DocTipo) + ' ' + rTrim(DocSerie)+Convert(VarChar(10), DocNumero) Documento," & _
                " RTrim(IsNull(CEmFantasia, rTrim(CPeApellido1) + ', ' + RTrim(CPeNombre1))) Cliente, ISNULL(CliCiRUC, '') [C.I./R.U.C.], NotDevuelve" & _
                " FROM Documento INNER JOIN Nota ON NotNota = DocCodigo INNER JOIN CLiente ON DocCliente = CliCodigo" & _
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
                chkNotaPendiente.Tag = rsAux("NotDevuelve")
            End If
            rsAux.Close
        
        ElseIf Val(hliCliente.Tag) > 0 Then
            
            'Cargo a partir de un cliente.
            'Si estoy buscando para cambio de productos entonces le pido que seleccione un documento.
            
            'Busco los documentos que tengan artículos entregados.
            'dbo.ListaArticulosDelDocumento(DocCodigo)
            Cons = "SELECT DocCodigo, DocFModificacion, DocTotal, DocFecha Fecha, dbo.NombreTipoDocumento(100+DocTipo) + ' ' + rTrim(DocSerie)+'-'+Convert(VarChar(10), DocNumero) Documento, rtrim(ArtNombre) Artículos " & _
                " FROM Documento INNER JOIN Devolucion ON DocCodigo = DevNota AND DevNota IS NOT NULL AND DevLocal IS NULL AND DevAnulada IS NULL " & _
                " INNER JOIN Renglon ON RenDocumento = DocCodigo " & _
                " INNER JOIN Articulo ON ArtID = RenArticulo AND ArtTipo <> 151" & _
                " WHERE DocTipo IN (3,4,10) AND DocCliente = " & hliCliente.Tag & " ORDER BY DocFecha DESC"
                
            Set objHelp = New clsListadeAyuda
            If objHelp.ActivarAyuda(cBase, Cons, 5500, 2, "Búsqueda") > 0 Then
                hliDocumento.Tag = objHelp.RetornoDatoSeleccionado(0)
                hliDocumento.Caption = objHelp.RetornoDatoSeleccionado(4)
                lblDocumento.Tag = objHelp.RetornoDatoSeleccionado(1)
                chkNotaPendiente.Tag = objHelp.RetornoDatoSeleccionado(2)
            Else
                MsgBox "No hay un documento para el cliente que tenga artículos a devolver.", vbInformation, "Cambio de producto"
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
        If Val(hliDocumento.Tag) > 0 Then
            
            If Val(hliCliente.Tag) > 0 Then
                BuscoComentariosAlerta Val(hliCliente.Tag), True
            End If
            
            If Val(hliDocumento.Tag) > 0 Then
                
                'Busco si el documento posee artículos disponibles para devolver.
                BuscarArticulosParaDevolverDelDocumento
                If lstArticulos.Rows = 1 Then
                    MsgBox "Atención el documento no posee artículos entregados o del mismo no se pueden devolver más artículos.", vbInformation, "ATENCIÓN"
                Else
                    butAceptar.Enabled = True
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


Private Sub txtEnvioCompra_Change()
    If Val(txtEnvioCompra.Tag) > 0 Then
        LimpiarControlesEnvio
    End If
End Sub

Private Sub txtEnvioCompra_GotFocus()
    MostrarAyuda
    With txtEnvioCompra
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtEnvioCompra_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(txtEnvioCompra.Tag) > 0 Then
            chkIncluirAlPendiente.SetFocus
        Else
            'busco el envío.
            BuscarEnvioCompra
        End If
    End If
End Sub

Private Sub BuscarEnvioCompra()
On Error GoTo errBE
    
    chkIncluirAlPendiente.value = 0
    chkIncluirAlPendiente.Enabled = False
            
    Cons = "SELECT *, RTrim(IsNull(CEmFantasia, rTrim(CPeApellido1) + ', ' + RTrim(CPeNombre1))) Cliente, ISNULL(CliCiRUC, '') [C.I./R.U.C.], CASE WHEN DocCodigo > 0 THEN dbo.NombreTipoDocumento(100+DocTipo) ELSE 'Vta.Telef.' END TipoDoc  " & _
        " FROM Envio " & _
        " LEFT OUTER JOIN Documento ON DocCodigo = EnvDocumento AND EnvTipo = 1 " & _
        " LEFT OUTER JOIN VentaTelefonica ON VTeCodigo = EnvDocumento AND EnvTipo = 3 " & _
        " INNER JOIN Cliente ON CliCodigo = EnvCliente " & _
        " LEFT OUTER JOIN CPersona ON CPeCliente = CliCodigo LEFT OUTER JOIN CEmpresa ON CEmCliente = CliCodigo " & _
        " WHERE EnvCodigo = " & Val(txtEnvioCompra.Text) & " AND EnvEstado = 0 " & _
        " AND EnvCodigo NOT IN (SELECT EReEnvio FROM EnviosRemitos)"
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
            
        hliDocEnvio.Tag = rsAux("EnvTipo") & "|" & rsAux("EnvDocumento")
        If rsAux("EnvTipo") = 3 Then
            hliDocEnvio.Caption = "Vta.Tel. " & rsAux("EnvDocumento")
        Else
            hliDocEnvio.Caption = rsAux("TipoDoc") + " " + rsAux("DocSerie") & "-" & rsAux("DocNumero")
            chkIncluirAlPendiente.value = 1
            chkIncluirAlPendiente.Enabled = True
            chkIncluirAlPendiente.Tag = rsAux("DocTipo")
        End If
        
        hliClienteEnvio.Caption = "(" & RTrim(rsAux("C.I./R.U.C.")) & ") " & rsAux("Cliente")
        hliClienteEnvio.Tag = rsAux("EnvCliente")
        txtEnvioCompra.Tag = rsAux("EnvCodigo")
        
        If (Val(hliClienteEnvio.Tag) <> Val(hliCliente.Tag)) Then
            MsgBox "El cliente del envío no coincide con el cliente de la nota, verifique.", vbInformation, "POSIBLE ERROR"
        End If
        
    Else
        
        txtEnvioCompra.Tag = ""
        MsgBox "No hay resultados para el dato ingresado.", vbInformation, "Búsqueda"
        
    End If
    butAceptar.Enabled = (Val(txtEnvioCompra.Tag) > 0)
    rsAux.Close
    
    Exit Sub
errBE:
    clsGeneral.OcurrioError "Error al buscar el envío.", Err.Description, "Error al buscar el envío"
    Screen.MousePointer = 0
End Sub

Private Sub txtEnvioCompra_LostFocus()
    MostrarAyuda ""
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


