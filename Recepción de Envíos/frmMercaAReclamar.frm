VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D851F632-A4E6-4F61-863C-9480B5EC86D9}#1.2#0"; "orgDateCtrl.ocx"
Object = "{162F4D73-979C-4E83-84D4-C9D8E6AB1FE3}#1.7#0"; "orgCtrlFlat.ocx"
Begin VB.Form frmMercaAReclamar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Retornar envío impreso"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picDatos 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1575
      ScaleWidth      =   7095
      TabIndex        =   12
      Top             =   480
      Width           =   7095
      Begin orgDateCtrl.orgDate tFecha 
         Height          =   315
         Left            =   3600
         TabIndex        =   3
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   39304
      End
      Begin OrgCtrlFlat.orgComboFlat cHora 
         Height          =   315
         Left            =   5520
         TabIndex        =   5
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         WidthListBox    =   0
      End
      Begin OrgCtrlFlat.orgComboFlat cbCombo 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         WidthListBox    =   0
      End
      Begin VB.TextBox tMotivo 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Text            =   "frmMercaAReclamar.frx":0000
         Top             =   1080
         Width           =   6975
      End
      Begin VB.CheckBox chSendMsg 
         Appearance      =   0  'Flat
         Caption         =   "Enviar mensaje"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1800
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lbTitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Condiciones para el nuevo estado"
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
         Left            =   120
         TabIndex        =   13
         Top             =   60
         Width           =   6735
      End
      Begin VB.Label lbfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lbMemo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Comentario:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lbCombo 
         BackStyle       =   0  'Transparent
         Caption         =   "&Estado:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lbHora 
         BackStyle       =   0  'Transparent
         Caption         =   "&Hora"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5040
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   540
      Left            =   2400
      TabIndex        =   11
      Top             =   4440
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   953
      ButtonWidth     =   2037
      ButtonHeight    =   953
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
            Key             =   "save"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Devuelve todo"
            Key             =   "devuelve"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Retiene todo"
            Key             =   "retiene"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "exit"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsArticulos 
      Height          =   2175
      Left            =   0
      TabIndex        =   9
      Top             =   2160
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3836
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
      SheetBorder     =   -2147483639
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   4
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   2
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5520
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMercaAReclamar.frx":0006
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMercaAReclamar.frx":0118
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMercaAReclamar.frx":046A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMercaAReclamar.frx":077C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   7215
      TabIndex        =   14
      Top             =   0
      Width           =   7215
      Begin OrgCtrlFlat.orgHiperLink hlVaCon 
         Height          =   255
         Left            =   1560
         Top             =   90
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
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
         ForeColorOver   =   8388608
         Caption         =   "Va Con"
         BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ArrowCaption    =   4
      End
      Begin VB.Label lbCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Envío: 8888888"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1290
      End
      Begin VB.Label lbDireccion 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Av Italia 2545/604"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   2520
         TabIndex        =   15
         Top             =   120
         Width           =   4455
      End
   End
   Begin VB.Label lbMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Dividir un envío que está en entrega se utiliza para dejar los artículos que no fueron entregados al cliente en un nuevo envío"
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
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   5280
      Width           =   6735
   End
   Begin VB.Shape shfac 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      FillColor       =   &H00DCFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   7020
   End
   Begin VB.Menu MnuVaCon 
      Caption         =   "VaCon"
      Visible         =   0   'False
      Begin VB.Menu MnuVaConItem 
         Caption         =   "item"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMercaAReclamar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iTipoDoc As Byte, iDocumento As Long, iCliente As Long, bCobraVta As Boolean
Private Type tDatosFlete
    Agenda As Double
    AgendaAbierta As Double
    AgendaCierre As Date
    HorarioRango As Integer
    HoraEnvio As String
End Type
Private rDatosFlete As tDatosFlete

Private lIDEnvioCobroVta As Long        '--> me indicá si es vta telefónica con cobro.
Private lAgeEnvio As Integer

Public prmInvocacion As Byte    '0) cambio fecha 1) anular el envío, 2) cambia camión
Public prmEnvio As Long

Private Function BuscoNombreMoneda(ByVal Codigo As Long, ByRef sSigno As String) As String
    On Error GoTo ErrBU
    Dim RsMoneda As rdoResultset
    BuscoNombreMoneda = ""
    Cons = "Select MonNombre, MonSigno FROM Moneda WHERE MonCodigo = " & Codigo
    Set RsMoneda = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsMoneda.EOF Then
        BuscoNombreMoneda = Trim(RsMoneda!MonNombre)
        sSigno = Trim(RsMoneda!MonSigno)
    End If
    RsMoneda.Close
    Exit Function
ErrBU:
End Function

Private Sub ImprimoNota(lnTipoDoc As Long, lnDocumento As Long)
Dim RsCr As rdoResultset
Dim strCliente As String, strRuc As String, strDireccion As String
Dim sSigno As String, sMoneda As String

Dim jobnum As Integer
Dim JobSRep1 As Integer
Dim JobSRep2 As Integer
Dim Result As Integer
Dim CantForm   As Integer
Dim NombreFormula As String

On Error GoTo ErrCrystal

    jobnum = crAbroReporte(gPathListados & "NotaDevolucion.RPT")
    If jobnum = 0 Then GoTo ErrCrystal
    
    'Configuro la Impresora
    
    If Trim(Printer.DeviceName) <> Trim(paDevNCtdo) Then SeteoImpresoraPorDefecto paDevNCtdo
    If Not crSeteoImpresora(jobnum, Printer, paBandNCtdo) Then GoTo ErrCrystal
        
    'Saco Datos del cliente.
    Cons = "Select * From Cliente" _
                & " LEFT OUTER JOIN CPersona ON CliCodigo = CPeCliente" _
                & " LEFT OUTER JOIN CEmpresa ON CliCodigo = CEmCliente" _
        & " Where CliCodigo = " & iCliente
    
    Set RsCr = cBase.OpenResultset(Cons, rdOpenDynamic)
           
    If RsCr!CliTipo = 2 Then
        If Not IsNull(RsCr!CEmNombre) Then
            strCliente = Trim(RsCr!CEmNombre)
        Else
            strCliente = Trim(RsCr!CEmFantasia)
        End If
        If Not IsNull(RsCr!CliCIRuc) Then
            strRuc = Trim(RsCr!CliCIRuc)
        Else
            strRuc = ""
        End If
    Else
        strCliente = Trim(Trim(Format(RsCr!CPeApellido1, "#")) & " " & Trim(Format(RsCr!CPeApellido2, "#"))) & ", " & Trim(Trim(Format(RsCr!CPeNombre1, "#")) & " " & Trim(Format(RsCr!CPeNombre2, "#")))
        If Not IsNull(RsCr!CliCIRuc) Then
            strCliente = strCliente & " (" & Trim(objGral.RetornoFormatoCedula(RsCr!CliCIRuc)) & ")"
        End If
        If Not IsNull(RsCr!CPERuc) Then
            strRuc = Trim(RsCr!CPERuc)
        Else
            strRuc = ""
        End If
    End If
    If Not IsNull(RsCr!CliDireccion) Then
        strDireccion = Trim(objGral.ArmoDireccionEnTexto(cBase, RsCr!CliDireccion))
    Else
        strDireccion = ""
    End If
    RsCr.Close
    
    sMoneda = BuscoNombreMoneda(1, sSigno)
    
    'Obtengo la cantidad de formulas que tiene el reporte.
    CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
    If CantForm = -1 Then GoTo ErrCrystal
    
    Dim I As Integer
    'Cargo Propiedades para el reporte Contado --------------------------------------------------------------------------------
    For I = 0 To CantForm - 1
        NombreFormula = crObtengoNombreFormula(jobnum, I)
        
        Select Case LCase(NombreFormula)
            Case "": GoTo ErrCrystal
            Case "nombredocumento"
                Result = crSeteoFormula(jobnum%, NombreFormula, "'" & paDNDevolucion & "'")
            Case "cliente":   Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(strCliente) & "'")
            Case "direccion": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(strDireccion) & "'")
            Case "ruc": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(objGral.RetornoFormatoRuc(strRuc)) & "'")
            Case "codigobarras"
                If lnTipoDoc = TipoDocumento.Contado Then
                    Result = crSeteoFormula(jobnum%, NombreFormula, "'" & CodigoDeBarras(TipoDocumento.NotaDevolucion, lnDocumento) & "'")
                Else
                    Result = crSeteoFormula(jobnum%, NombreFormula, "'" & CodigoDeBarras(TipoDocumento.NotaCredito, lnDocumento) & "'")
                End If
            
            Case "signomoneda": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(sSigno) & "'")
            Case "nombremoneda": Result = crSeteoFormula(jobnum%, NombreFormula, "'(" & sMoneda & ")'")
    
            Case "textoretira"
                'Detallamos el documento al cual se le hace la nota.
                'Result = crSeteoFormula(jobnum%, NombreFormula, "'" & lFactura.Caption & "'")
            
            Case Else: Result = 1
        End Select
        If Result = 0 Then GoTo ErrCrystal
    Next
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    'Seteo la Query del reporte-----------------------------------------------------------------
    Cons = "SELECT Documento.DocCodigo , Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal, Documento.DocIVA, Documento.DocVendedor" _
            & " From " & paBD & ".dbo.Documento Documento " _
            & " Where DocCodigo = " & lnDocumento
    If crSeteoSqlQuery(jobnum%, Cons) = 0 Then GoTo ErrCrystal
        
    'Subreporte srContado.rpt  y srContado.rpt - 01-----------------------------------------------------------------------------
    JobSRep1 = crAbroSubreporte(jobnum, "srContado.rpt")
    If JobSRep1 = 0 Then GoTo ErrCrystal
    
    Cons = "SELECT Renglon.RenDocumento, Renglon.RenCantidad, Renglon.RenPrecio, Renglon.RenDescripcion," _
            & " From { oj " & paBD & ".dbo.Renglon Renglon INNER JOIN " _
                           & paBD & ".dbo.Articulo Articulo ON Renglon.RenArticulo = Articulo.ArtId}"
    If crSeteoSqlQuery(JobSRep1, Cons) = 0 Then GoTo ErrCrystal
    
    JobSRep2 = crAbroSubreporte(jobnum, "srContado.rpt - 01")
    If JobSRep2 = 0 Then GoTo ErrCrystal
    If crSeteoSqlQuery(JobSRep2, Cons) = 0 Then GoTo ErrCrystal
    '-------------------------------------------------------------------------------------------------------------------------------------
    
    
    If crMandoAImpresora(jobnum, 1) = 0 Then GoTo ErrCrystal
    If Not crInicioImpresion(jobnum, True, False) Then GoTo ErrCrystal
    
    If Not crCierroSubReporte(JobSRep1) Then GoTo ErrCrystal
    If Not crCierroSubReporte(JobSRep2) Then GoTo ErrCrystal
    crCierroTrabajo jobnum
    Screen.MousePointer = 0
    Exit Sub

ErrCrystal:
    Screen.MousePointer = 0
    objGral.OcurrioError crMsgErr
    On Error Resume Next
    Screen.MousePointer = 11
    crCierroSubReporte JobSRep1
    crCierroSubReporte JobSRep2
    Screen.MousePointer = 0
    Exit Sub

End Sub

Private Function f_HayDiferenciasConDocumento() As Currency
Dim rsD As rdoResultset
    f_HayDiferenciasConDocumento = 0
    'Busco todas las diferencias de Envío que tenga cobro en algún documento que no haya sido anulado
    'o que no tenga nota.
    Cons = "Select IsNull(Sum(DevValorFlete), 0) From DiferenciaEnvio, Documento" & _
                " Where DEvEnvio = " & Me.prmEnvio & _
                " And DEvDocumento = DocCodigo And DocAnulado = 0 " & _
                " And DEvDocumento Not In (Select NotFactura From Nota)"
    Set rsD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    f_HayDiferenciasConDocumento = rsD(0)
    rsD.Close
End Function

Private Function f_ShowSuceso(ByVal sTitulo As String, ByRef sDefensa As String) As Long
    'Retorno el id del usuario y la defensa x referencia
Dim objSuceso As New clsSuceso
    
    objSuceso.ActivoFormulario paCodigoDeUsuario, sTitulo, cBase
    f_ShowSuceso = objSuceso.RetornoValor(True)
    sDefensa = objSuceso.RetornoValor(False, True)
    Set objSuceso = Nothing
    Me.Refresh
    
End Function

Private Sub loc_SaveAnular()
Dim lNota As Long
Dim sDefensa As String: sDefensa = ""
Dim lUID As Long
Dim bAnularRem As Boolean
On Error GoTo errSA
    'Controlo diferencias de envíos.
    
    If f_HayDiferenciasConDocumento Then
        MsgBox "Existen diferencias de envío facturadas." & vbCr & vbCr & "Primero debe eliminarlas, acceda al formulario de envíos para hacerlo.", vbExclamation, "Atención"
        Exit Sub
    End If
    
        
    If bCobraVta Then
        '........................................................................................................
        '           Este envío es el que cobra la venta telefónica.
        '........................................................................................................
        
        'Veo el caso 5
        'Válido que el documento no tenga nota.
        
        If iTipoDoc = 6 Then
            Dim iDocC As Long
    'OJO ESTO NUNCA OCURRE YA QUE NO DEJO ELIMINAR UN ENVIO QUE TIENE UN VACON
            'La vta está en un remito x lo que está en un va con o es única.
            If hlVaCon.Visible Then
                Cons = "SELECT EVCDocumento FROM EnvioVaCon WHERE EVCEnvio = " & prmEnvio
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                iDocC = RsAux("EVCDocumento")
                RsAux.Close

            Else
                Cons = "SELECT RDoDocumento FROM RemitoDocumento WHERE RDoRemito = " & iDocumento
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                iDocC = RsAux("RDoDocumento")
                RsAux.Close
            
            End If
            
            Cons = "SELECT EnvCodigo FROM Envio " & _
                    " WHERE EnvCodigo <> " & prmEnvio & _
                    " AND (((EnvCodigo IN(SELECT EnvCodigo FROM Envio WHERE EnvDocumento = " & iDocC & " And EnvCodigo <> " & prmEnvio & "))" & _
                    " OR EnvCodigo IN (SELECT EnvCodigo FROM Envio, RemitoDocumento WHERE RDoRemito <> " & iDocumento & " And RDoDocumento = " & iDocC & " And RDoRemito = EnvDocumento And EnvTipo = 1))" & _
                    " OR EnvCodigo IN (SELECT EVCEnvio From EnvioVaCon WHERE EVCDocumento = " & iDocC & " And EVCEnvio <> " & prmEnvio & "))"
            
            
        Else
            Cons = "Select EnvCodigo From Envio" & _
                " Where EnvCodigo <> " & prmEnvio & _
                " And (EnvDocumento = " & iDocumento & _
                " Or EnvDocumento IN(Select RDoRemito From RemitoDocumento Where RDoDocumento = " & iDocumento & "))"
            
        End If
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If Not RsAux.EOF Then
            RsAux.Close
            'Primero veo si hay otros envíos para el documento
            MsgBox "Este envío tiene el cobro de una venta telefónica, antes debe eliminar los otros envíos del documento.", vbInformation, "ATENCIÓN"
            Exit Sub
        End If
        RsAux.Close
    End If
    
    
    'No son vtas sin facturar
    'Verifico si el envío tiene asociado documentospendientes --> método nuevo
    'Si es así doy aviso de lo que hago con ellos y procedo a eliminar.
    Cons = "Select * From DocumentoPendiente, Documento " & _
        "Where DPeTipo = 1 And DPeIDTipo = " & Me.prmEnvio & " And DPeDocumento = DocCodigo And DocAnulado = 0 And DPeIDLiquidacion Is Null"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        MsgBox "El envío posee al menos un documento impreso en depósito, los mismos serán anulados o se les haran la nota correspondiente.", vbInformation, "Documentos Pendientes"
    End If
    RsAux.Close
    
    sDefensa = "¿Confirma eliminar el envío seleccionado?"
    If iTipoDoc = 6 Then
        Cons = "Select * from RemitoDocumento Where RDoRemito = " & iDocumento
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            bAnularRem = True
            sDefensa = "El envío pertenece a un remito impreso en depósito al eliminar el envío se anulará el remito." & vbCrLf & IIf(hlVaCon.Visible, "NO SE ELIMINA EL VA CON", "") & vbCrLf & vbCrLf & sDefensa
        End If
        RsAux.Close
    End If
    
    If MsgBox(sDefensa, vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
        sDefensa = ""
        FechaDelServidor
            
        'Llamo al registro del Suceso-------------------------------------------------------------
        lUID = f_ShowSuceso("Eliminación de Envíos", sDefensa)
        If lUID = 0 Then Screen.MousePointer = 0: Exit Sub
    
        Dim sXML As String
        sXML = fnc_GetQArticulos
    
        On Error GoTo errBT
        cBase.BeginTrans
        On Error GoTo Resumo
        
        Cons = "EXEC prg_Envio_Eliminar " & prmEnvio & ", '" & Format(CDate(vsArticulos.Tag), "yyyy/mm/dd hh:nn:ss") & "', " & lUID & ", " & paCodigoDeSucursal & ", '" & sXML & "', " & paCodigoDeTerminal
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux(0) = -1 Then
            Cons = RsAux(1)
            RsAux.Close
            cBase.RollbackTrans
            MsgBox "No se logró eliminar el envío." & vbCrLf & vbCrLf & "Detalle: " & Cons, vbCritical, "Atención"
            Exit Sub
        Else
            'Saco las notas.
            Cons = ""
            If Not IsNull(RsAux(1)) Then Cons = Trim(RsAux(1))
        End If
        RsAux.Close
        
        objGral.RegistroSuceso cBase, gFechaServidor, 5, paCodigoDeTerminal, lUID, 0, Descripcion:="Envío Nº " & prmEnvio, Defensa:=Trim(sDefensa)
        
        cBase.CommitTrans
        
        On Error Resume Next
        If Cons <> "" Then
            'hay notas para imprimir
            Dim arrNota() As String
            arrNota = Split(Cons, ",")
            If UBound(arrNota) > 0 Then
                For lUID = 1 To UBound(arrNota)
                    If Val(arrNota(lUID)) > 0 Then ImprimoNota TipoDocumento.Contado, Val(arrNota(lUID))
                Next
            End If
        End If
        Unload Me
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
errSA:
    Screen.MousePointer = vbDefault
    objGral.OcurrioError "Error inesperado al intentar anular.", Err.Description
    Exit Sub
    
errBT:
    Screen.MousePointer = vbDefault
    objGral.OcurrioError "Error inesperado al iniciar la trasacción.", Err.Description
    Exit Sub
    
Relajo:
    On Error Resume Next
    RsAux.Close
    cBase.RollbackTrans
    Screen.MousePointer = vbDefault
    objGral.OcurrioError "Error inesperado intentar eliminar el envío.", Err.Description
    Exit Sub
    
Resumo:
    Resume Relajo

End Sub

Private Sub loc_EnvioConDocumentoPendiente()
On Error GoTo errEDP
Dim rsD As rdoResultset
Dim sQy As String
    sQy = "Select DocSerie, DocNumero From DocumentoPendiente, Documento " & _
         "WHERE DPeTipo = 1 " & _
         "AND DPeIDTipo IN (" & _
                "SELECT EnvCodigo FROM Envio " & _
                " WHERE (EnvCodigo = " & Me.prmEnvio & " OR EnvCodigo IN(SELECT EVCEnvio FROM EnvioVaCon WHERE EVCID = (SELECT EVCID FROM EnvioVaCon WHERE EVCEnvio = " & Me.prmEnvio & ")))" & _
                " And EnvEstado = 3)" & _
         " And DPeDocumento = DocCodigo"
    
    Set rsD = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    sQy = ""
    Do While Not rsD.EOF
        sQy = sQy & IIf(sQy = "", "", ", ") & rsD("DocSerie") & " " & rsD("DocNumero")
        rsD.MoveNext
    Loop
    rsD.Close
    If sQy <> "" Then
        MsgBox "Atención el envío seleccionado posee los siguientes documentos pendientes asociados a él y debe reclamarselos al camionero si aún los posee: " & vbCrLf & sQy, vbInformation, "Facturas del envío"
    End If
    Exit Sub
errEDP:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al buscar si el documento posee documentos pendientes.", Err.Description, "Documentos pendientes"
End Sub

Private Function fnc_GetQArticulos() As String
'Armo el xml
Dim sXML As String, sRenglon As String
    '(ArtID  int, QEnvio smallint, QDevuelve int, QPendiente int, FEdit DateTime)
    sRenglon = "<Renglon ArtID=""[mIDArt]"" QEnvio=""[mQEnvio]"" QDevuelve=""[mQDevuelve]"" QPendiente=""[mQPendiente]"" FEdit=""[mFEdit]""></Renglon>"
Dim iQ As Integer
    With vsArticulos
        For iQ = 1 To .Rows - 1
            Cons = Replace(sRenglon, "[mIDArt]", .Cell(flexcpData, iQ, 0))
            Cons = Replace(Cons, "[mQEnvio]", .Cell(flexcpData, iQ, 1))
            Cons = Replace(Cons, "[mQDevuelve]", .Cell(flexcpText, iQ, 5))
            Cons = Replace(Cons, "[mQPendiente]", .Cell(flexcpText, iQ, 4))
            Cons = Replace(Cons, "[mFEdit]", Format(CDate(.Cell(flexcpData, iQ, 4)), "yyyy/mm/dd hh:nn:ss"))
            sXML = sXML & Cons
        Next
        sXML = "<ROOT>" & sXML & "</ROOT>"
    End With
    fnc_GetQArticulos = sXML
End Function

Private Function fnc_DBSaveNuevoEstado() As Boolean
Dim iEstadoEnvio As Byte
Dim iNewCamion As Integer
    
    On Error GoTo errInit
    If Me.prmInvocacion = 0 Then
        If cbCombo.ListIndex = 0 Then iEstadoEnvio = EstadoEnvio.AConfirmar Else iEstadoEnvio = EstadoEnvio.AImprimir
    ElseIf Me.prmInvocacion = 2 Then
        'Cambio de camión
        iEstadoEnvio = EstadoEnvio.AImprimir
        iNewCamion = cbCombo.ItemData(cbCombo.ListIndex)
    Else
    
    End If
    
    Screen.MousePointer = 11
    'Aquí tengo que controlar el documento pendiente
    'Si es con el formato anterior doy msg y me voy.
    
    '10/8/2007 quedamos en que solo aviso.
    'If Not fnc_ValidoNoEntregado() Then Screen.MousePointer = 0: Exit Function
    fnc_MensajeVentaTelefonica
    
    'Doy aviso si el envío tiene una factura pendiente.
    loc_EnvioConDocumentoPendiente
    
    'Verifico si hay envíos dentro del remito que fueron entregados
    'Si es así anulo el remito y le retorno la mercadería al documento
    If iTipoDoc = 6 Then
        'Si un envío lo dividio entonces no es un va con OJO este caso es excepcional.
        Cons = "SELECT EnvCodigo FROM Envio WHERE EnvDocumento = " & iDocumento & _
            "AND EnvCodigo <> " & prmEnvio & " And EnvEstado = 4"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            MsgBox "Atención este envío pertenece a un remito que ya posee envíos entregados." & vbCrLf & vbCrLf & "El remito será anulado.", vbInformation, "Atención"
        End If
        RsAux.Close
    End If
    
        
'Armo la hora
Dim sHora As String, sFecha As String
    If tFecha.HasValue Then
        If Trim(cHora.Text) <> vbNullString Then
            If cHora.ListIndex > -1 Then
                'Busco en codigotexto el valor.
                If cHora.ItemData(cHora.ListIndex) > 0 Then
                    Cons = "Select * from CodigoTexto Where Codigo = " & cHora.ItemData(cHora.ListIndex)
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If Len(Trim(RsAux!Clase)) < 4 Then
                        sHora = "0" & Trim(RsAux!Clase) & "-" & Trim(RsAux!Puntaje)
                    Else
                        sHora = Trim(RsAux!Clase) & "-" & Trim(RsAux!Puntaje)
                    End If
                    RsAux.Close
                Else
                    sHora = cHora.Text
                End If
            Else
                sHora = Trim(cHora.Text)
            End If
        End If
        sFecha = Format(tFecha.Value, "yyyy/mm/dd")
    Else
        iEstadoEnvio = EstadoEnvio.AConfirmar
    End If

    Dim sXML As String
    sXML = fnc_GetQArticulos

    cBase.BeginTrans
    On Error GoTo errRB
    Cons = "EXEC prg_RecepcionEnvio_RetornoConNuevoEstado " & Me.prmEnvio & ", '" & sFecha & "', '" & sHora & "', " & _
                    iEstadoEnvio & ", '" & Trim(tMotivo.Text) & "', " & paTComEnvConf & ", " & paCodigoDeUsuario & ", '" & sXML & "', " & paCodigoDeSucursal & ", " & paCodigoDeTerminal & ", " & iNewCamion
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux(0) = 1 Then
        cBase.CommitTrans
        On Error Resume Next
        If iEstadoEnvio = 0 And chSendMsg.Value Then
            Dim objConnect As New clsConexion
            objConnect.EnviaMensaje paUIDEnvConf, "Envío(s) a Confirmar", tMotivo.Text, DateAdd("s", 30, Now), 751, paCodigoDeUsuario
            Set objConnect = Nothing
        End If
        MsgBox "Envío actualizado.", vbInformation, "Grabar"
        fnc_DBSaveNuevoEstado = True
    Else
        If Not IsNull(RsAux(1)) Then Cons = Trim(RsAux(1)) Else Cons = ""
        RsAux.Close
        cBase.RollbackTrans
        MsgBox "No se logró grabar la información, refresque la información y vuelva a ingresar los datos." & vbCrLf & vbCrLf & "Detalle: " & Cons, vbExclamation, "Error al grabar"
    End If
    Screen.MousePointer = 0
    RsAux.Close
    Exit Function
    
errInit:
    Screen.MousePointer = vbDefault
    objGral.OcurrioError "Error al armar la información a almacenar.", Err.Description
    Exit Function
    
errSave:
    Screen.MousePointer = vbDefault
    cBase.RollbackTrans
    objGral.OcurrioError "Error al grabar la información.", Err.Description
    Exit Function
    
errRB:
    Resume errSave
    Exit Function
End Function

Private Function fnc_MensajeVentaTelefonica() As Boolean
Dim rsV As rdoResultset
'Si tengo una vta telefónica con más de un envío --> si o si se hizo remito.

    'Busco si tengo otro envío en la vta telefónica que este en otro remito.
    If lIDEnvioCobroVta > 0 Then
        If iTipoDoc = 6 Then
            'La vta está en un remito x lo que está en un va con o es única.
            If hlVaCon.Visible Then
                'Está dentro de un Va Con
                Cons = "Select Count(*) From Envio" & _
                    " Where EnvCodigo <> " & lIDEnvioCobroVta & _
                    " And EXISTS( " & _
                        "SELECT * FROM RemitoDocumento, VentaTelefonica " & _
                        "WHERE RDoDocumento IN (SELECT EVCDocumento From EnvioVaCon Where EVCEnvio = " & lIDEnvioCobroVta & ")" & _
                        "And RDoDocumento = VTeDocumento And (EnvDocumento = RDoDocumento or EnvDocumento = RDoRemito))"

            Else
                Cons = "Select Count(*) From Envio" & _
                    " Where EnvCodigo <> " & lIDEnvioCobroVta & _
                    " And ((EnvDocumento IN(Select RDoDocumento From RemitoDocumento Where RDoRemito = " & iDocumento & "))" & _
                    " Or (EnvDocumento IN(Select RDoRemito From RemitoDocumento Where RDoRemito = " & iDocumento & ")))"
            End If
        Else
            Cons = "Select EnvCodigo From Envio" & _
                " Where EnvCodigo <> " & lIDEnvioCobroVta & _
                " And (EnvDocumento = " & iDocumento & _
                " Or EnvDocumento IN(Select RDoRemito From RemitoDocumento Where RDoDocumento = " & iDocumento & "))"
        End If
        
'        Cons = "SELECT RDoRemito " & _
            "FROM RemitoDocumento R1 " & _
            "WHERE EXISTS (SELECT * FROM Envio, RemitoDocumento R2, VentaTelefonica " & _
                    "WHERE EnvCodigo = " & prmEnvio & " And EnvDocumento = R2.RDoRemito And R2.RDoDocumento = VTeDocumento " & _
                    "AND R1.RDoRemito <> R2.RDoRemito And R1.RDoDocumento = VTeDocumento)"
        Set rsV = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Cons = ""
        If Not rsV.EOF Then
            If Not IsNull(rsV(0)) Then
                If rsV(0) > 1 Then
                    If lIDEnvioCobroVta <> Me.prmEnvio Then
                        Cons = "El envío " & lIDEnvioCobroVta & " pertenece al Va Con y el mismo cobra una venta telefónica que posee más de un envío." & vbCrLf & vbCrLf _
                            & "Controle el estado de los otros envíos."
                    Else
                        Cons = "El envío cobra una venta telefónica que posee más de un envío." & vbCrLf & vbCrLf _
                            & "Controle el estado de los otros envíos."
                    End If
                    MsgBox Cons, vbInformation, "Atención"
                End If
            End If
        End If
        rsV.Close
        fnc_MensajeVentaTelefonica = (Cons = "")
    End If
End Function


Private Function fnc_ValidateGrabar() As Boolean
    
    If Me.prmInvocacion <> 1 Then
        If cbCombo.ListIndex = -1 And cbCombo.Text <> "" Then
            MsgBox "El dato ingresado en el combo no es correcto.", vbExclamation, "Validación"
            cbCombo.SetFocus
            Exit Function
        End If
        If tFecha.Visible Then
            If tFecha.HasValue Then
                If CDate(tFecha.Text) < Date Then
                    MsgBox "La fecha es menor a hoy.", vbExclamation, "Atención"
                    tFecha.SetFocus
                    Exit Function
                End If
                If Not f_EsDiaAbierto Then
                    If MsgBox("El día no está abierto." & vbCr & vbCr & "¿Confirma guardar el envío con esa fecha?", vbQuestion + vbYesNo + vbDefaultButton2, "Posible Error") = vbNo Then Exit Function
                End If
                If Not ValidoRangoHorario Then Exit Function
            Else
                If cbCombo.ListIndex = 1 Or Me.prmInvocacion = 2 Then
                    MsgBox "Debe indicar la nueva fecha de envío.", vbExclamation, "Atención"
                    tFecha.SetFocus
                    Exit Function
                End If
            End If
        End If
        If Trim(tMotivo.Text) = "" And cbCombo.ListIndex = 0 Then
            MsgBox "Ingrese el motivo por el cual pone el envío a confirmar.", vbExclamation, "Atención"
            tMotivo.SetFocus
            Exit Function
        End If
    End If
    fnc_ValidateGrabar = True
End Function

Private Sub loc_StateMotivo()
    If cbCombo.ListIndex = 0 Then
        tMotivo.Enabled = True:: tMotivo.BackColor = vbWindowBackground
        chSendMsg.Enabled = True
    Else
        tMotivo.Enabled = False: tMotivo.Text = "": tMotivo.BackColor = vbButtonFace
        chSendMsg.Enabled = False
    End If
End Sub

Private Function ValidoRangoHorario() As Boolean

    ValidoRangoHorario = True
    If cHora.ListIndex > -1 Then Exit Function
    
    If InStr(1, cHora.Text, "-") > 0 Then
        Select Case Len(cHora.Text)
            Case 9
                If CLng(Mid(cHora.Text, 1, InStr(1, cHora.Text, "-") - 1)) > CLng(Mid(cHora.Text, InStr(1, cHora.Text, "-") + 1, Len(cHora.Text))) Then
                    MsgBox "El rango de horario ingresado no es válido.", vbExclamation, "ATENCIÓN"
                    cHora.SetFocus
                    ValidoRangoHorario = False
                    Exit Function
                End If
                
            Case 5
                If InStr(1, cHora.Text, "-") = 1 Then
                    If CLng(Mid(cHora.Text, InStr(1, cHora.Text, "-") + 1, Len(cHora.Text))) < paPrimeraHoraEnvio Then
                        MsgBox "El horario ingresado es menor a la primera hora de entrega.", vbExclamation, "ATENCIÓN"
                        ValidoRangoHorario = False
                        Exit Function
                    Else
                        If paPrimeraHoraEnvio < 1000 Then
                            cHora.Text = "0" & paPrimeraHoraEnvio & cHora.Text
                        Else
                            cHora.Text = paPrimeraHoraEnvio & cHora.Text
                        End If
                        Exit Function
                    End If
                Else
                    If InStr(1, cHora.Text, "-") = 5 Then
                        If CLng(Mid(cHora.Text, 1, InStr(1, cHora.Text, "-") - 1)) > paUltimaHoraEnvio Then
                            MsgBox "El horario ingresado es mayor que la última hora de envio.", vbExclamation, "ATENCIÓN"
                            ValidoRangoHorario = False
                            Exit Function
                        Else
                            cHora.Text = cHora.Text & paUltimaHoraEnvio
                        End If
                    Else
                        MsgBox "No se ingreso un horario válido. [####-####]", vbExclamation, "ATENCIÓN"
                        cHora.SetFocus
                        ValidoRangoHorario = False
                        Exit Function
                    End If
                End If
            
            Case 8
                If CLng(Mid(cHora.Text, 1, InStr(1, cHora.Text, "-") - 1)) > CLng(Mid(cHora.Text, InStr(1, cHora.Text, "-") + 1, Len(cHora.Text))) Then
                    MsgBox "El rango de horario ingresado no es válido.", vbExclamation, "ATENCIÓN"
                    cHora.SetFocus
                    ValidoRangoHorario = False
                    Exit Function
                End If
                
                If InStr(1, cHora.Text, "-") = 4 Then
                    cHora.Text = "0" & cHora.Text
                End If
            
            Case Else
                    MsgBox "No se ingreso un horario válido. [####-####]", vbExclamation, "ATENCIÓN"
                    cHora.SetFocus
                    ValidoRangoHorario = False
                    Exit Function
                    
        End Select
    Else
        MsgBox "No se ingreso un horario válido. [####-####]", vbExclamation, "ATENCIÓN"
        cHora.SetFocus
        ValidoRangoHorario = False
        Exit Function
    End If
    
    'Ahora válido el rango de horas.
    If Val(tFecha.Tag) > 0 And rDatosFlete.HorarioRango > 0 Then
        
        Dim lhora As Long
        
        lhora = (CLng(Mid(cHora.Text, InStr(1, cHora.Text, "-") + 1, Len(cHora.Text))) - CLng(Mid(cHora.Text, 1, InStr(1, cHora.Text, "-") - 1))) / 100
        If lhora < rDatosFlete.HorarioRango Then
            If MsgBox("El rango ingresado es menor al posible para el flete seleccionado." & vbCr & vbCr & _
                        "El flete tiene un rango de " & rDatosFlete.HorarioRango & " hora(s) y se asigno un rango de " & lhora & " hora(s)" & vbCr & vbCr & _
                        "¿Confirma mantener el rango ingresado?", vbQuestion + vbYesNo + vbDefaultButton2, "Posible error en horario") = vbNo Then
                cHora.SetFocus
                ValidoRangoHorario = False
            End If
        End If
    End If
    
End Function

Private Function BuscoProximoDia(dFecha As Date, strMat As String)
Dim rsHora As rdoResultset
Dim intDia As Integer, intSuma As Integer
    
    'Por las dudas que no cumpla en la semana paso la agenda normal.
    
    On Error GoTo errBDER
    
    BuscoProximoDia = -1
    
    'Consulto en base a la matriz devuelta.
    Cons = "Select * From HorarioFlete Where HFlIndice IN (" & strMat & ")"
    Set rsHora = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsHora.EOF Then
        
        'Busco el valor que coincida con el dia de hoy y ahí busco para arriba.
        intSuma = 0
        Do While intSuma < 7
            rsHora.MoveFirst
            intDia = Weekday(dFecha + intSuma)
            Do While Not rsHora.EOF
                If rsHora!HFlDiaSemana = intDia Then
                    BuscoProximoDia = intSuma
                    GoTo Encontre
                End If
                rsHora.MoveNext
            Loop
            intSuma = intSuma + 1
        Loop
        rsHora.Close
    End If

Encontre:
    rsHora.Close
    Exit Function
    
errBDER:
    objGral.OcurrioError "Error al buscar el primer día disponible para el tipo de flete.", Trim(Err.Description)
End Function

Private Function db_FindZona(lCodDireccion As Long) As Long
On Error GoTo errFZ
Dim lZonP As Long
Dim lIDComp As Long

    Cons = "Select IsNull(CZoZona,0) as CZoZona, IsNull(DirComplejo,0) as DirComplejo From Direccion " _
            & " Left Outer Join CalleZona On DirCalle = CZoCalle And CZoDesde <= DirPuerta And CZoHasta >= DirPuerta" _
        & " Where DirCodigo = " & lCodDireccion
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        lZonP = 0
        lIDComp = 0
    Else
        lIDComp = RsAux!DirComplejo
        lZonP = RsAux!CZoZona
    End If
    RsAux.Close
    
    If lIDComp > 0 Then
        'Si tengo complejo --> busco la zona para el mismo.
        Cons = "Select CZoZona From Complejo, CalleZona" _
            & " Where ComCodigo = " & lIDComp _
            & " And CZoCalle = ComCalle And CZoDesde <= ComNumero And CZoHasta >= ComNumero"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            lZonP = RsAux!CZoZona
        End If
        RsAux.Close
    End If
    db_FindZona = lZonP
    Exit Function

errFZ:
    objGral.OcurrioError "Error al buscar el código de la zona.", Err.Description
End Function

Private Sub s_SetHoraEntrega()
On Error GoTo errCHEPD
Dim sMat As String
    Screen.MousePointer = 11
    cHora.Clear
    If DateDiff("d", rDatosFlete.AgendaCierre, Date) >= 7 Then
        'Como cerro hace una semana tomo la agenda normal.
        sMat = superp_MatrizSuperposicion(rDatosFlete.Agenda)
    Else
        sMat = superp_MatrizSuperposicion(rDatosFlete.AgendaAbierta)
    End If
    If rDatosFlete.HoraEnvio <> "" Then
        loc_SetHoraEnvio rDatosFlete.HoraEnvio, sMat
    Else
        If sMat <> "" Then
            Cons = "Select HFlCodigo, HFlNombre From HorarioFlete Where HFlIndice IN (" & sMat & ")" _
                & " And HFlDiaSemana = " & Weekday(CDate(tFecha.Value)) & " Order by HFlInicio"
            CargoCombo Cons, cHora
        End If
    End If
    If cHora.ListCount > 0 Then cHora.ListIndex = 0
    Screen.MousePointer = 0
    Exit Sub
errCHEPD:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al buscar los horarios para el día de semana.", Trim(Err.Description)
End Sub
Private Sub loc_SetHoraEnvio(ByVal sHora As String, ByVal sMat As String)
Dim arrHoraE() As String, arrID() As String
Dim iQ As Integer
On Error Resume Next
Dim rsHF As rdoResultset
Dim sIn As String

    arrHoraE = Split(sHora, ",")
    
    Cons = "Select HEnIndice From HorarioFlete, HoraEnvio Where HFlIndice IN (" & sMat & ")" _
            & " And HFlDiaSemana = " & Weekday(CDate(tFecha.Value)) _
            & " And HEnCodigo = HFlCodigo  Order by HFlInicio"
    Set rsHF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsHF.EOF
        If sIn <> "" Then sIn = sIn & ","
        sIn = sIn & Trim(rsHF("HEnIndice"))
        rsHF.MoveNext
    Loop
    rsHF.Close
    If sIn <> "" Then sIn = "," & sIn & ","
    
    For iQ = 0 To UBound(arrHoraE)
        arrID = Split(arrHoraE(iQ), ":")
        If InStr(1, sIn, "," & arrID(0) & ",") > 0 Then cHora.AddItem arrID(1)
    Next
End Sub

Private Function f_EsDiaAbierto() As Boolean
On Error GoTo errEA
Dim sMat As String
Dim dAux As Date

    f_EsDiaAbierto = False
    If Val(tFecha.Tag) = 0 Then Exit Function
    
    If DateDiff("d", rDatosFlete.AgendaCierre, Date) >= 7 Then
        'Como cerro hace una semana tomo la agenda normal.
        sMat = superp_MatrizSuperposicion(rDatosFlete.Agenda)
    Else
        sMat = superp_MatrizSuperposicion(rDatosFlete.AgendaAbierta)
    End If
    dAux = tFecha.Value
    If sMat <> "" Then f_EsDiaAbierto = (BuscoProximoDia(dAux, sMat) = 0)
errEA:
End Function

Private Sub s_SetFirstDay()
Dim sMat As String
Dim iSuma As Integer
Dim dAux As Date
    
    If rDatosFlete.AgendaCierre < Date Then dAux = Date Else dAux = rDatosFlete.AgendaCierre
    
    If DateDiff("d", rDatosFlete.AgendaCierre, Date) >= 7 Then
        'Como cerro hace una semana tomo la agenda normal.
        sMat = superp_MatrizSuperposicion(rDatosFlete.Agenda)
    Else
        sMat = superp_MatrizSuperposicion(rDatosFlete.AgendaAbierta)
    End If
    tFecha.Clear
    If sMat <> "" Then
        iSuma = BuscoProximoDia(dAux, sMat)
        If iSuma <> -1 Then tFecha.Value = Format(DateAdd("d", iSuma, dAux), "dd/mm/yyyy")
    End If
    If Not tFecha.HasValue Then
        MsgBox "No hay agenda abierta para el tipo de flete del envío.", vbExclamation, "Atención"
        tFecha.Value = Date
    End If
    
End Sub

Private Sub s_GetDatosTipoFlete()
On Error GoTo errGD
Dim RsF As rdoResultset
Dim lZona As Long

    With rDatosFlete
        .Agenda = 0
        .AgendaAbierta = 0
        .AgendaCierre = Date
        .HoraEnvio = ""
        .HorarioRango = 0
    End With
    
    'Ya lo cargue o no hay tipo de flete
    If Val(tFecha.Tag) = 0 Or rDatosFlete.Agenda > 0 Then Exit Sub
    Screen.MousePointer = 11
    Cons = "Select IsNull(TFlAgenda, 0) as Agenda, IsNull(TFlAgendaHabilitada, 0) as AgendaH, IsNull(TFLFechaAgeHab, GetDate()) as FAgenda, TFLHoraEnvio, IsNull(THoRangoHS, 0) as RangoHS " & _
                " From TipoFlete " & _
                        "Left Outer Join TipoHorario On TFlRangoHs = THoID" & _
                " Where TFLCodigo = " & Val(tFecha.Tag)
    Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsF.EOF Then
        With rDatosFlete
            .Agenda = RsF("Agenda")
            .AgendaAbierta = RsF("AgendaH")
            .AgendaCierre = RsF("FAgenda")
            If Not IsNull(RsF("TFLHoraEnvio")) Then .HoraEnvio = Trim(RsF!TFLHoraEnvio)
            .HorarioRango = RsF("RangoHS")
        End With
    End If
    RsF.Close
    
    'Si no es de Agencia --> busco para la zona.
    If lAgeEnvio > 0 Then
        'Tengo que buscar la zona de la agencia.
        Cons = "Select IsNull(CZoZona, 0) From Agencia, Direccion, CalleZona" _
                & " Where AgeCodigo = " & lAgeEnvio _
                & " And AgeDireccion = DirCodigo And DirCalle = CZoCalle " _
                & " And CZoDesde <= DirPuerta And CZoHasta >= DirPuerta"

        Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsF.EOF Then
            lZona = RsF(0)
        End If
        RsF.Close
    End If

    'Si no tengo zona de agencia busco para la dirección del envío.
    If lZona = 0 Then lZona = db_FindZona(Val(lbDireccion.Tag))
        
    Cons = "Select * From FleteAgendaZona " & _
            " Where FAZZona = " & lZona & " And FAZTipoFlete = " & Val(tFecha.Tag)
    Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsF.EOF Then
        With rDatosFlete
            If Not IsNull(RsF!FAZAgenda) Then
                .Agenda = RsF!FAZAgenda
                If Not IsNull(RsF!FAZAgendaHabilitada) Then .AgendaAbierta = RsF!FAZAgendaHabilitada Else .AgendaAbierta = .Agenda
                If Not IsNull(RsF!FAZFechaAgeHab) Then .AgendaCierre = RsF("FAZFechaAgeHab")
            End If
            If Not IsNull(RsF!FAZRangoHS) Then .HorarioRango = RsF!FAZRangoHS
            If Not IsNull(RsF!FAZHoraEnvio) Then .HoraEnvio = Trim(RsF!FAZHoraEnvio)
        End With
    End If
    RsF.Close
    Screen.MousePointer = 0
    Exit Sub
errGD:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al buscar los datos del tipo de flete.", Err.Description, "Datos tipo de flete"
End Sub

Private Sub loc_SetGridDevRet(ByVal bDevuelve As Boolean)
Dim iQ As Integer
Dim iCol As Byte, iCol0 As Byte
    If bDevuelve Then iCol0 = 4: iCol = 5 Else iCol = 4: iCol0 = 5
    
    With vsArticulos
        For iQ = .FixedRows To .Rows - 1
            .Cell(flexcpBackColor, iQ, 4, , 5) = vbWindowBackground
            If Val(.Cell(flexcpText, iQ, iCol0)) > 0 Then
                .Cell(flexcpText, iQ, iCol) = Val(.Cell(flexcpText, iQ, iCol)) + Val(.Cell(flexcpText, iQ, iCol0))
                .Cell(flexcpText, iQ, iCol0) = 0
            End If
            If Val(.Cell(flexcpText, iQ, iCol)) > 0 Then .Cell(flexcpBackColor, iQ, iCol) = &HADDEFF '&H66CCFF
        Next
    End With
End Sub
Private Sub loc_SetColorNormal(ByVal bDevuelve As Boolean)
Dim iQ As Integer
Dim iCol As Integer
    If bDevuelve Then iCol = 5 Else iCol = 4
    With vsArticulos
        For iQ = .FixedRows To .Rows - 1
            If .Cell(flexcpBackColor, iQ, iCol) = &HADDEFF Then
                .Cell(flexcpBackColor, iQ, iCol) = vbWindowBackground
            End If
        Next
    End With
End Sub

Private Sub loc_DBDevuelvePendiente(ByVal bDevuelve As Boolean)
On Error GoTo errInit
    If Not fnc_ValidateGrabar Then Exit Sub
    
    Screen.MousePointer = 11
    'pongo todo como devuelto en la lista.
    loc_SetGridDevRet bDevuelve
    'pregunto
    Screen.MousePointer = 0
    If Me.prmInvocacion = 1 Then
        loc_SaveAnular
        Exit Sub
    End If
    
    
    If bDevuelve Then
        Cons = "¿Confirma grabar la información?" & vbCrLf & vbCrLf & "El camionero DEVUELVE LA MERCADERÍA DEL ENVÍO" & vbCrLf & "Si desea puede validar en la grilla las cantidades ajustadas."
    Else
        Cons = "¿Confirma grabar la información?" & vbCrLf & vbCrLf & "El camionero RETIENE LA MERCADERÍA DEL ENVÍO" & vbCrLf & "Si desea puede validar en la grilla las cantidades ajustadas."
    End If
    If MsgBox(Cons, vbQuestion + vbYesNo, "Grabar") = vbYes Then
        If fnc_DBSaveNuevoEstado Then Unload Me
    Else
        loc_SetColorNormal True
    End If
    Exit Sub
errInit:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al armar la grilla para grabar.", Err.Description, "Grabar"
End Sub

Private Sub loc_FindEnvio()
On Error GoTo errFE
Dim lAux As Long
Dim iCodImpresion As Integer
Dim sCodEnvios As String

    Screen.MousePointer = 11
    Toolbar1.Buttons("save").Enabled = False
    vsArticulos.Rows = 1
    tFecha.Clear
    
    MnuVaConItem(0).Tag = ""
    
    Cons = "Select EnvCodigo, EnvAgencia, EnvTipoFlete, EnvEstado, EnvFModificacion, EnvDireccion, EnvCodImpresion, EnvCamion, " & _
            "EnvFechaPrometida, EnvRangoHora, EnvReclamoCobro, EnvDocumento, EnvCliente, IsNull(DocTipo, 0) DT" & _
        " FROM (Envio LEFT OUTER JOIN Documento ON EnvDocumento = DocCodigo) " & _
        "WHERE (EnvCodigo = " & Me.prmEnvio & " OR EnvCodigo IN(SELECT EVCEnvio FROM EnvioVaCon WHERE EVCID = (SELECT EVCID FROM EnvioVaCon WHERE EVCEnvio = " & Me.prmEnvio & "))) " & _
        "AND EnvCodImpresion Is Not Null"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        RsAux.Close
        Screen.MousePointer = 0
        MsgBox "No existe un envío con ese código.", vbExclamation, "Atención"
        Exit Sub
    Else
        
        iCodImpresion = RsAux("EnvCodImpresion")
        iTipoDoc = RsAux("DT")
        iDocumento = RsAux("EnvDocumento")
        
        If RsAux("EnvEstado") <> 3 Then
            Screen.MousePointer = 0
            RsAux.Close
            MsgBox "El envío no tiene el estado impreso, para modificarlo acceda al formulario de envíos.", vbExclamation, "Atención"
            Exit Sub
        Else
            Do While Not RsAux.EOF
                
                If Not IsNull(RsAux("EnvReclamoCobro")) And lIDEnvioCobroVta = 0 Then lIDEnvioCobroVta = RsAux("EnvCodigo")
                
                
                sCodEnvios = sCodEnvios & IIf(sCodEnvios = "", "", ", ") & RsAux("EnvCodigo")
            
                If RsAux("EnvCodigo") = Me.prmEnvio Then
                    bCobraVta = Not IsNull(RsAux("EnvReclamoCobro"))
                    iCliente = RsAux("EnvCliente")
                    lbDireccion.Caption = objGral.ArmoDireccionEnTexto(cBase, RsAux("EnvDireccion"))
                    lbDireccion.Tag = RsAux("EnvDireccion")
                    vsArticulos.Tag = RsAux("EnvFModificacion")
                    tFecha.Tag = RsAux!EnvTipoFlete
                    If Not IsNull(RsAux("EnvCamion")) And Me.prmInvocacion = 2 Then cbCombo.ListIndex = cbCombo.FindItemData(RsAux("EnvCamion"))
                    If Not IsNull(RsAux!EnvAgencia) Then lAgeEnvio = RsAux!EnvAgencia
                    
                    If Not IsNull(RsAux("EnvFechaPrometida")) Then
                        tFecha.Value = RsAux("EnvFechaPrometida")
                        If Not IsNull(RsAux("EnvRangoHora")) Then cHora.Text = Trim(RsAux("EnvRangoHora"))
                    End If
                    
                ElseIf RsAux("EnvCodigo") <> Me.prmEnvio Then
                    
                    If Val(MnuVaConItem(0).Tag) > 0 Then Load MnuVaConItem(MnuVaConItem.UBound + 1)
                    With MnuVaConItem(MnuVaConItem.UBound)
                        .Visible = True
                        .Enabled = True
                        .Caption = Trim(RsAux("EnvCodigo"))
                        .Tag = .Caption
                    End With
                    hlVaCon.Visible = True
                End If
                RsAux.MoveNext
            Loop
        End If
        RsAux.Close
    End If
        
    
    Cons = "Select Sum(REvAEntregar) as QArt, ReECantidadTotal as QT, ReECantidadEntregada as QE, ArtID, ArtCodigo, rTrim(ArtNombre) as ArtNombre, ReEFModificacion" & _
        " From RenglonEnvio, Articulo, RenglonEntrega " & _
        " Where REvEnvio IN (" & sCodEnvios & ")" & _
        " And RevArticulo = ArtID And RevAEntregar > 0 And ReEArticulo = ArtID And ReECodImpresion = " & iCodImpresion & _
        " Group by ArtID, ArtCodigo, ArtNombre, ReECantidadTotal, ReECantidadEntregada, ReEFModificacion"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    'Cargo la lista por si selecciona la opción EntregaParcial.
    Do While Not RsAux.EOF
        With vsArticulos
            .AddItem "(" & Format(RsAux!ArtCodigo, "000,000") & ") " & Trim(RsAux!ArtNombre)
            .Cell(flexcpText, .Rows - 1, 1) = RsAux("QArt")
            .Cell(flexcpText, .Rows - 1, 2) = RsAux("QT")
            .Cell(flexcpText, .Rows - 1, 3) = RsAux("QE")
            If RsAux("QE") = RsAux("QT") Or RsAux("QE") = 0 Then
                'El camión tiene o no tienen toda la mercadería por lo tanto devuelve todo
                If RsAux("QE") > 0 Then
                    .Cell(flexcpText, .Rows - 1, 4) = 0
                    .Cell(flexcpText, .Rows - 1, 5) = RsAux("QArt")
                Else
                    .Cell(flexcpText, .Rows - 1, 4) = 0
                    .Cell(flexcpText, .Rows - 1, 5) = 0
                    .Cell(flexcpBackColor, .Rows - 1, 4, , 5) = &HE0E0E0
                End If
            Else
                'El camión tiene asignada parte de la mercadería.
                'Por lo tanto siempre le voy a restar al camión.
                .Cell(flexcpText, .Rows - 1, 4) = 0
                If RsAux("QE") > RsAux("QArt") Then
                    .Cell(flexcpText, .Rows - 1, 5) = RsAux("QArt")
                Else
                    .Cell(flexcpText, .Rows - 1, 5) = RsAux("QE")
                End If
            End If
            .Cell(flexcpBackColor, .Rows - 1, 0, , 3) = vbWindowBackground
            .Cell(flexcpBackColor, .Rows - 1, 1) = &HFFF5F0 '14857624
            .Cell(flexcpBackColor, .Rows - 1, 3) = &HFFF5F0

            lAux = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = lAux
            lAux = RsAux("QArt"): .Cell(flexcpData, .Rows - 1, 1) = lAux
            lAux = RsAux("QT"): .Cell(flexcpData, .Rows - 1, 2) = lAux
            lAux = RsAux("QE"): .Cell(flexcpData, .Rows - 1, 3) = lAux
            Cons = RsAux("ReEFModificacion"): .Cell(flexcpData, .Rows - 1, 4) = Cons
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    Toolbar1.Buttons("save").Enabled = (vsArticulos.Rows > 1)
    
    hlVaCon.Visible = (MnuVaConItem(0).Tag <> "")
    
    On Error Resume Next
    If vsArticulos.Rows > 1 Then vsArticulos.SetFocus
    Screen.MousePointer = 0
    Exit Sub
errFE:
    Screen.MousePointer = 0
    vsArticulos.Rows = 1
    objGral.OcurrioError "Error al buscar el envío.", Err.Description
End Sub

Private Sub loc_DBSave()
    
    If Me.prmInvocacion <> 1 Then
        If MsgBox("¿Confirma grabar la información ingresada?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
            If fnc_DBSaveNuevoEstado Then Unload Me
        End If
    Else
        loc_SaveAnular
    End If
    Exit Sub
    
End Sub

Private Sub cbCombo_GotFocus()
    If prmInvocacion = 2 Then
        lbMsg.Caption = "Seleccione el nuevo camión para el envío."
    Else
        lbMsg.Caption = "Seleccione si el envío quedá asignado a una nueva fecha para el mismo camionero o lo pasa a confirmar"
    End If
End Sub

Private Sub cbCombo_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If prmInvocacion = 0 Then
            loc_StateMotivo
            tFecha.SetFocus
        Else
            vsArticulos.SetFocus
        End If
    End If
End Sub

Private Sub cbCombo_LostFocus()
    lbMsg.Caption = ""
End Sub

Private Sub cbCombo_Validate(Cancel As Boolean)
    loc_StateMotivo
End Sub

Private Sub cHora_GotFocus()
On Error Resume Next
    lbMsg.Caption = "Seleccione el horario a enviar o ingrese un rango con el formato ####-####"
End Sub

Private Sub cHora_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(cHora.Text) <> "" Then If Not ValidoRangoHorario Then Exit Sub
        If cbCombo.ListIndex = 0 Then
            If tMotivo.Enabled Then tMotivo.SetFocus
        Else
            vsArticulos.SetFocus
        End If
    End If
End Sub


Private Sub chSendMsg_GotFocus()
    lbMsg.Caption = "Indique si envía un mensaje con el motivo ingresado indicando que el envío cambio el estado."
End Sub

Private Sub chSendMsg_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then tMotivo.SetFocus
End Sub

Private Sub chSendMsg_LostFocus()
    lbMsg.Caption = ""
End Sub

Private Sub Form_Load()
On Error Resume Next
    picDatos.Visible = (Me.prmInvocacion <> 1)
    Select Case prmInvocacion
        Case 0
            With cbCombo
                .Clear
                .AddItem "A confirmar"
                .AddItem "Nueva fecha"
                .ListIndex = 1
            End With
            loc_StateMotivo
            lbCombo.Caption = "&Estado:"
            Me.Height = 6345
            Me.Caption = "Cambiar fecha a envío"
            cHora.Clear
            tMotivo.Text = ""
            
        Case 1
            vsArticulos.Top = picDatos.Top
            Me.Height = 4750
            picDatos.Height = 0
            Me.Caption = "Anular envío"
        
        Case 2
            Me.Caption = "Cambiar camión"
            Me.Height = 6345 - 840
            picDatos.Height = 735
            lbTitulo.Caption = "Cambio de camionero"
            lbCombo.Caption = "&Camión:"
            CargoCombo "Select CamCodigo, CamNombre From Camion Order By CamNombre", cbCombo
    End Select
    
    vsArticulos.Top = picDatos.Top + picDatos.Height + 120
    Toolbar1.Top = vsArticulos.Top + vsArticulos.Height + 120
    shfac.Top = Toolbar1.Top + Toolbar1.Height + 120
    lbMsg.Top = shfac.Top + 120
    
    With vsArticulos
        .Rows = 1
        .FixedRows = 1
        .FormatString = "Artículo|Q Env|Q CImp|Entregada|Retiene|>Devuelve"
        .FixedCols = 4
        .RowHeight(0) = 315
        .ColWidth(0) = 3400
        .BackColorSel = vbInfoBackground
        .ForeColorSel = vbWindowText
        .FocusRect = flexFocusHeavy
    End With
    Toolbar1.Buttons("save").Enabled = False
    lbCodigo.Caption = "Envío: " & Me.prmEnvio
    
    hlVaCon.Left = lbCodigo.Left + lbCodigo.Width + 120
    lbDireccion.Left = hlVaCon.Left + 960
    loc_FindEnvio
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    vsArticulos.Left = 60
    vsArticulos.Width = ScaleWidth - 120
End Sub

Private Sub hlVaCon_Click()
    PopupMenu MnuVaCon, , hlVaCon.Left, hlVaCon.Top + hlVaCon.Height
End Sub

Private Sub tFecha_Change()
    cHora.Clear
End Sub

Private Sub tFecha_GotFocus()
    With tFecha
        If Not .HasValue Then
            s_GetDatosTipoFlete
            s_SetFirstDay
            s_SetHoraEntrega
        End If
    End With
    lbMsg.Caption = "Ingrese la fecha en que se volverá a enviar."
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If tFecha.HasValue Then
            If Not f_EsDiaAbierto Then
                MsgBox "El día ingresado no está abierto.", vbExclamation, "Atención"
            Else
                s_SetHoraEntrega
                cHora.SetFocus
            End If
        End If
    End If
End Sub

Private Sub tMotivo_GotFocus()
    lbMsg.Caption = "Ingrese el motivo por el cual el envío queda a confirmar, se graba un comentario para el cliente."
End Sub

Private Sub tMotivo_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn And Trim(tMotivo.Text) <> "" Then vsArticulos.SetFocus
End Sub

Private Sub tMotivo_LostFocus()
    lbMsg.Caption = ""
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "save": loc_DBSave
        Case "devuelve": loc_DBDevuelvePendiente True
        Case "retiene": loc_DBDevuelvePendiente False
        Case "exit": Unload Me
    End Select
    Screen.MousePointer = 0
End Sub

Private Sub vsArticulos_GotFocus()
    lbMsg.Caption = "Seleccione la columna e ingrese la cantidad de artículos que retiene o devuelve el camionero. (+ o - suma o resta). Las filas en gris no puede modificarlas."
End Sub

Private Sub vsArticulos_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errkd
Dim iC1 As Byte, iC2 As Byte
    If Shift <> 0 Then Exit Sub
    With vsArticulos
        If .Cell(flexcpBackColor, .Row, 5) = &HE0E0E0 Then Exit Sub
        Select Case KeyCode
            Case vbKeyAdd
                'Dada la columna que es resto a la otra.
                iC1 = .Col
                If .Col = 5 Then iC2 = 4 Else iC2 = 5
                If Val(.Cell(flexcpText, .Row, iC2)) > 0 Then
                    .Cell(flexcpText, .Row, iC1) = Val(.Cell(flexcpText, .Row, iC1)) + 1
                    .Cell(flexcpText, .Row, iC2) = Val(.Cell(flexcpText, .Row, iC2)) - 1
                End If
            Case vbKeySubtract
                'Dada la columna que es sumo a la otra.
                iC1 = .Col
                If .Col = 5 Then iC2 = 4 Else iC2 = 5
                If Val(.Cell(flexcpText, .Row, iC1)) > 0 Then
                    .Cell(flexcpText, .Row, iC1) = Val(.Cell(flexcpText, .Row, iC1)) - 1
                    .Cell(flexcpText, .Row, iC2) = Val(.Cell(flexcpText, .Row, iC2)) + 1
                End If
        End Select
    End With
errkd:
End Sub

Private Sub vsArticulos_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn And Toolbar1.Buttons("save").Enabled Then loc_DBSave
End Sub
