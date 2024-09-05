VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{191D08B9-4E92-4372-BF17-417911F14390}#1.5#0"; "orGridPreview.ocx"
Begin VB.Form frmControl 
   Appearance      =   0  'Flat
   Caption         =   "Notas Sin Facturar"
   ClientHeight    =   6735
   ClientLeft      =   2790
   ClientTop       =   1995
   ClientWidth     =   6330
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmControl.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   6330
   Begin VB.CommandButton bPrint 
      Caption         =   "Imprimir"
      Height          =   315
      Left            =   3900
      TabIndex        =   5
      Top             =   60
      Width           =   1095
   End
   Begin VB.CommandButton bGrabar 
      Caption         =   "Hacer Nota"
      Height          =   315
      Left            =   5100
      TabIndex        =   4
      Top             =   60
      Width           =   1095
   End
   Begin VB.CommandButton bCargar 
      Caption         =   "&Cargar Lista"
      Height          =   315
      Left            =   2100
      TabIndex        =   2
      Top             =   60
      Width           =   1095
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
      Height          =   5115
      Left            =   60
      TabIndex        =   3
      Top             =   540
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9022
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
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
      BackColorFixed  =   -2147483645
      ForeColorFixed  =   -2147483634
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14737632
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   8421631
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   4
      GridLinesFixed  =   5
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
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
   Begin MSComCtl2.DTPicker tFecha 
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Top             =   60
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   3997697
      CurrentDate     =   37543
   End
   Begin orGridPreview.GridPreview cPrint 
      Left            =   3480
      Top             =   420
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lFecha 
      BackStyle       =   0  'Transparent
      Caption         =   "&Fecha:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   615
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "?"
      Visible         =   0   'False
      Begin VB.Menu MnuHlp 
         Caption         =   "Ayuda"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim prmListado As Boolean

Private Enum cols
    Numero
    MoraTotal
    MoraIVA
    Estado
End Enum

Private prmFecha As Date
Private ProdInteresMora As clsProducto

Private Sub bCargar_Click()
    fnc_CargoListas
End Sub

Private Sub bGrabar_Click()
    If Not prmListado Then
        MsgBox "Para emitir la nota primero debe realizar la impresión del listado.", vbInformation, "Falta Listar"
        Exit Sub
    End If
    AccionGrabar
End Sub

Private Sub bPrint_Click()
    
    With cPrint
        .Orientation = opPortrait
        .Caption = Me.Caption
        .Header = Me.Caption & " al " & tFecha.Value
        .PageBorder = opTopBottom
        .MarginLeft = 1200
        .MarginTop = 800
    End With
    
    Screen.MousePointer = 11
    vsLista.ExtendLastCol = False
    
    With cPrint
        .Columns = 2
        .AddGrid vsLista.hwnd
        .LineAfterGrid ""
        .LineAfterGrid "Norma resolución 1.790 del 2006, art 44 decreto 597/88 con la redacción dada por el art. 1 del Decreto 388/92"

        .ShowPreview
    End With
    
    vsLista.ExtendLastCol = True
    
    prmListado = True
    Screen.MousePointer = 0

End Sub

Private Sub Form_Load()
    
    If Not ValidarVersionEFactura Then
        MsgBox "La versión del componente CGSAEFactura está desactualizado, debe distribuir software." _
                    & vbCrLf & vbCrLf & "Se cancelará la ejecución.", vbCritical, "EFactura"
        End
    End If
    InicializoForm
    CargoArticuloInteresesPorMora
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    On Error GoTo errRZ
    vsLista.Height = Me.ScaleHeight - vsLista.Top - 100
errRZ:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EndMain
End Sub

Private Sub InicializoForm()
    
    On Error Resume Next
    LimpioFicha
   
    With vsLista
        .Rows = 1: .cols = 1
        .FormatString = "<Recibo|>Mora Total|>IVA|"
        .ColWidth(cols.Numero) = 1100
        .ColWidth(cols.MoraTotal) = 1400
        .ColWidth(cols.MoraIVA) = 1200
        .RowHeight(0) = 260
        .WordWrap = False
        .MergeCells = flexMergeSpill
        .ExtendLastCol = True
        
        .Editable = False
        .RowHeight(0) = 280
        .SelectionMode = flexSelectionByRow
    End With

End Sub

Private Function fnc_CargoListas()
On Error GoTo errCL
Dim xValor As Long, xQTotal As Integer
Dim xMoraTotal As Currency, xMoraIVA As Currency

    Screen.MousePointer = 11
    xQTotal = 0
    vsLista.Rows = vsLista.FixedRows
    prmFecha = tFecha.Value
    
    mSQL = "Select * from Documento, NotasSinFacturar" & _
                " Where DocCodigo = NSFIDRecibo " & _
                " And DocFecha Between '" & Format(prmFecha, "mm/dd/yyyy 00:00:00") & "' And '" & Format(prmFecha, "mm/dd/yyyy 23:59:59") & "'" & _
                " And NSFEstado = 0"
                
    Set RsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        Do While Not RsAux.EOF
            With vsLista
               
                .AddItem ""
                xValor = RsAux!DocCodigo: .Cell(flexcpData, .Rows - 1, cols.Numero) = xValor
                .Cell(flexcpText, .Rows - 1, cols.Numero) = Trim(RsAux!DocSerie) & "-" & RsAux!DocNumero
                
                .Cell(flexcpText, .Rows - 1, cols.MoraTotal) = Format(RsAux!NSFImporteTotal, "#,##0.00")
                .Cell(flexcpText, .Rows - 1, cols.MoraIVA) = Format(RsAux!NSFIVA, "#,##0.00")
                
                xQTotal = xQTotal + 1
                xMoraTotal = xMoraTotal + .Cell(flexcpValue, .Rows - 1, cols.MoraTotal)
                xMoraIVA = xMoraIVA + .Cell(flexcpValue, .Rows - 1, cols.MoraIVA)
            End With
            RsAux.MoveNext
        Loop
        
        With vsLista
            .AddItem "": .AddItem ""

            .Cell(flexcpText, .Rows - 1, cols.Numero) = xQTotal
            .Cell(flexcpText, .Rows - 1, cols.MoraTotal) = Format(xMoraTotal, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, cols.MoraIVA) = Format(xMoraIVA, "#,##0.00")
            .Cell(flexcpFontBold, .Rows - 1, 0, , .cols - 1) = True
        End With
    End If
    RsAux.Close
    Screen.MousePointer = 0
    
    If (vsLista.Rows = vsLista.FixedRows) Then
        MsgBox "No hay Notas de Débito para generar al " & Format(prmFecha, "dd/mm/yyyy") & ".", vbInformation, "No hay datos"
        Exit Function
    End If
    '----------------------    ----------------------   ----------------------  ----------------------  ----------------------
    bGrabar.Enabled = (prmFecha = Date)
    prmListado = False
    Exit Function
errCL:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar la lista de documentos.", Err.Description
End Function

Private Function ValidoGrabar() As Boolean
On Error GoTo errValidar

    ValidoGrabar = False
    If prmFecha <> Date Then
        MsgBox "Ud no puede emitir una nota de débito anterior al día de hoy.", vbInformation, "Posible Error"
        Exit Function
    End If
    ValidoGrabar = True
    
errValidar:
End Function

Private Sub AccionGrabar()
    
    On Error GoTo errorBT
   
    If Not ValidoGrabar Then Exit Sub
    
    If MsgBox("¿Confirma generar y emitir la Nota de Débito pendiente?", vbQuestion + vbYesNo, "Generar Nota Débito") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    Dim Serie As String, Numero As Long, txtRET As String, mTotalMora As Currency, mIVA As Currency
    Dim xIDNotaDebito As Long, xIDMoneda As Integer

    FechaDelServidor
   
    cBase.BeginTrans    'COMIENZO TRANSACCION ------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
    
    '1) Saco el total de la nota y el iva   ----------------------------
    Dim mTotal As Currency, mMora As Currency
    
    mSQL = "Select DocMoneda, Sum(NSFImporteTotal) as Total, Sum(NSFIva) as IVA from Documento, NotasSinFacturar" & _
                " Where DocCodigo = NSFIDRecibo " & _
                " And DocFecha Between '" & Format(prmFecha, "mm/dd/yyyy 00:00:00") & "' And '" & Format(prmFecha, "mm/dd/yyyy 23:59:59") & "'" & _
                " And DocAnulado = 0 " & _
                " And NSFEstado = 0" & _
                " Group By DocMoneda"
                
    Set RsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux("Total")) Then mTotalMora = RsAux("Total")
        If Not IsNull(RsAux("IVA")) Then mIVA = RsAux("IVA")
        If Not IsNull(RsAux("DocMoneda")) Then xIDMoneda = RsAux("DocMoneda")
    End If
    RsAux.Close
    
    If mTotalMora = 0 Then Err.Raise 1001, , "Nota ya generada"
        
    Dim CAE As clsCAEDocumento
    Set CAE = New clsCAEDocumento
    If Val(prmEFacturaProductivo) = 0 Then
        txtRET = NumeroDocumento(paDNDebito)
        Serie = Trim(Mid(txtRET, 1, 1))
        Numero = CLng(Trim(Mid(txtRET, 2, Len(txtRET))))
        With CAE
            .Desde = 1
            .Hasta = 9999999
            .Serie = Serie
            .Numero = Numero
            .IdDGI = "9014113"
            .TipoCFE = CFE_eFacturaNotaDebito
            .Vencimiento = "31/12/" & CStr(Year(Date))
        End With
    Else
        Dim caeG As New clsCAEGenerador
        Set CAE = caeG.ObtenerNumeroCAEDocumento(cBase, CFE_eFacturaNotaDebito, paCodigoDeSucursal)
        Set caeG = Nothing
    End If
       
    Cons = "Select * from Documento Where DocCodigo = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.AddNew
    RsAux!DocFecha = gFechaServidor
    RsAux!DocTipo = TipoDocumento.NotaDebito
    RsAux!DocSerie = Trim(Serie)
    RsAux!DocNumero = Numero
    RsAux!DocCliente = 1 'CGSA
    RsAux!DocMoneda = xIDMoneda
    RsAux!DocTotal = mTotalMora
    RsAux!DocIVA = mIVA

    RsAux!DocAnulado = 0
    RsAux!DocSucursal = paCodigoDeSucursal
    RsAux!DocUsuario = paCodigoDeUsuario
    RsAux!DocFModificacion = gFechaServidor
    
    RsAux.Update
    RsAux.Close
    '--------------------------------------------------------------------------------------------------------------------------------------------

    Cons = "SELECT MAX(DocCodigo) From Documento" & _
                " WHERE DocTipo = " & TipoDocumento.NotaDebito & _
                " AND DocSerie = '" & Serie & "'" & " AND DocNumero = " & Numero
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    xIDNotaDebito = RsAux(0)
    RsAux.Close
    

    Dim oNotaDeb As New clsDocumentoCGSA
    With oNotaDeb
        Set .Cliente = EmpresaEmisora
        .Codigo = xIDNotaDebito
        .Digitador = paCodigoDeUsuario
        .Emision = gFechaServidor
        .IVA = mIVA
        .Moneda.Codigo = xIDMoneda
        .Numero = CAE.Numero
        .Serie = CAE.Serie
        .sucursal = paCodigoDeSucursal
        .Tipo = TD_NotaDebito
        .Total = mTotalMora
    End With
    Dim oDocRel As clsDocumentoAsociado
    Dim oRenglon As New clsDocumentoRenglon
    With oRenglon
        .Cantidad = 1
        .IVA = mIVA
        .Precio = mTotalMora
        Set .Articulo = ProdInteresMora
    End With
    oNotaDeb.Renglones.Add oRenglon

'    mSQL = "Select NotasSinFacturar.* from Documento, NotasSinFacturar" & _
'                " Where DocCodigo = NSFIDRecibo " & _
'                " And DocFecha Between '" & Format(prmFecha, "mm/dd/yyyy 00:00:00") & "' And '" & Format(prmFecha, "mm/dd/yyyy 23:59:59") & "'" & _
'                " And DocAnulado = 0 " & _
'                " And NSFEstado = 0"

    mSQL = "SELECT NotasSinFacturar.*, Factura.DocTipo, Factura.DocFecha, Factura.DocCodigo, Factura.DocSerie, Factura.DocNumero, IsNull(EcomTipo, 0) EcomTipo " & _
        "FROM Documento Rec INNER JOIN NotasSinFacturar ON Rec.DocCodigo = NSFIDRecibo AND NSFEstado = 0 " & _
        "INNER JOIN DocumentoPago ON DPaDocQSalda = Rec.DocCodigo INNER JOIN Documento Factura ON Factura.DocCodigo = DPaDocASaldar " & _
        "LEFT OUTER JOIN eComprobantes ON EComID = Factura.DocCodigo " & _
        "WHERE Rec.DocFecha Between '" & Format(prmFecha, "mm/dd/yyyy 00:00:00") & "' And '" & Format(prmFecha, "mm/dd/yyyy 23:59:59") & "'" & _
        " AND Rec.DocAnulado = 0 ORDER BY NSFIDRecibo"

    Dim idRecAnt As Long
    Set RsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        If idRecAnt <> RsAux("NSFIDRecibo") Then
            '1) Relacion Recibo-Nueva Nota de débito
            mSQL = "Insert into DocumentoPago (DPaDocASaldar, DPaDocQSalda, DPaCuota, DPaDe, DPaAmortizacion, DPaMora)" & _
                        " Values (" & xIDNotaDebito & ", " & RsAux("NSFIDRecibo").Value & ", 1, 1, 0, 0)"
            
            cBase.Execute mSQL
        
            '2) Updateo el estado de la nota pendiente a procesada =1
            mSQL = "Update NotasSinFacturar Set NSFEstado = 1 Where NSFEstado = 0 And NSFIDRecibo =" & RsAux("NSFIDRecibo").Value
            cBase.Execute mSQL
            
            idRecAnt = RsAux("NSFIDRecibo")
        End If
        Set oDocRel = New clsDocumentoAsociado
        oNotaDeb.DocumentosAsociados.Add oDocRel
        With oDocRel
            .ID = RsAux("DocCodigo")
            .Fecha = RsAux("DocFecha")
            .Serie = Trim(RsAux("DocSerie"))
            .Numero = RsAux("DocNumero")
            .Tipo = RsAux("DocTipo")
            .TipoEFactura = RsAux("EcomTipo")
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If EmitirCFE(oNotaDeb, CAE) <> "" Then RsAux.Close: RsAux.Edit
    
    cBase.CommitTrans    'FIN DE TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

    On Error GoTo errFin
    If xIDNotaDebito > 0 Then
        crAbroEngine
        ImprimoNotaMoras xIDNotaDebito, xIDMoneda, TOPrinter:=True
        crCierroEngine
    End If
    
    Screen.MousePointer = 0
    
    
    LimpioFicha
    Exit Sub

    Exit Sub

errFin:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al restaurar el formulario, el documento fue emitido.", Err.Description
    Exit Sub

errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
    Exit Sub

errorET:
    Resume ErrorRoll
    
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación."
End Sub

Private Sub tFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then bCargar.SetFocus
End Sub

Private Sub vsLista_DblClick()
    On Error Resume Next
'    If vsLista.Rows > vsLista.FixedRows Then
'    End If
End Sub

Private Function LimpioFicha()

    tFecha.Value = Date
    vsLista.Rows = vsLista.FixedRows
    
    bGrabar.Enabled = False
        
End Function


Private Function ImprimoNotaMoras(idDocumento As Long, idMoneda As Integer, Optional TOPrinter As Boolean = True)
Dim job_Moras As Integer, job_R As Integer, job_QFrms As Integer, job_FrmName As String

Dim mMonNombre As String, mMonSigno As String, mConcepto As String

    On Error GoTo ErrCrystal
    fnc_DatosMoneda idMoneda, mMonSigno, mMonNombre
    
    job_Moras = crAbroReporte(prmPathListados & "Aporte.RPT")
    If job_Moras = 0 Then GoTo ErrCrystal
        
    job_QFrms = crObtengoCantidadFormulasEnReporte(job_Moras)
    If job_QFrms = -1 Then GoTo ErrCrystal

    If ChangeCnfgPrint Then     'Valido Si cambio la cfg de las impresoras      ------------------------------------
        prj_LoadConfigPrint bShowFrm:=False

        'Configuro la Impresora
        If Trim(Printer.DeviceName) <> Trim(paIReciboN) Then SeteoImpresoraPorDefecto paIReciboN
        'If Not crSeteoImpresora(job_Moras, Printer, paIReciboB, mOrientation:=2) Then GoTo ErrCrystal
        If Not crSeteoImpresora(job_Moras, Printer, paIReciboB, mOrientation:=2, paperSize:=13) Then GoTo ErrCrystal
    End If      '-------------------------------------------------------------------------------------------------------------
    
    'Dim arrWork() As String, arrValor() As String, idx As Integer

    
   
    Screen.MousePointer = 11
    
    'Cargo Propiedades para el reporte Contado --------------------------------
    For I = 0 To job_QFrms - 1
        job_FrmName = crObtengoNombreFormula(job_Moras, I)
        
        Select Case LCase(job_FrmName)
            Case "": GoTo ErrCrystal
            Case "nombredocumento": job_R = crSeteoFormula(job_Moras, job_FrmName, "'" & paDNDebito & "'")
            
            Case "cliente": job_R = crSeteoFormula(job_Moras%, job_FrmName, "''")
            Case "cedula": job_R = crSeteoFormula(job_Moras%, job_FrmName, "''")
                
            Case "ruc": job_R = crSeteoFormula(job_Moras%, job_FrmName, "''")
            
            Case "signomoneda": job_R = crSeteoFormula(job_Moras%, job_FrmName, "'" & mMonSigno & "'")
            Case "nombremoneda": job_R = crSeteoFormula(job_Moras%, job_FrmName, "'(" & mMonNombre & ")'")
                
            Case "usuario": job_R = crSeteoFormula(job_Moras%, job_FrmName, "'" & BuscoDigitoUsuario(paCodigoDeUsuario) & "'")
            
            Case "cuenta":
                    
                    mConcepto = "Norma resolución 1.790 del 2006, art 44 decreto 597/88 con la redacción dada por el art. 1 del Decreto 388/92"
                    job_R = crSeteoFormula(job_Moras%, job_FrmName, "'" & mConcepto & "'")
        
            Case "articulo":
                mConcepto = UCase("Concepto: Intereses por Mora")
                job_R = crSeteoFormula(job_Moras%, job_FrmName, "'" & mConcepto & "'")

            Case Else: job_R = 1
        End Select
        If job_R = 0 Then GoTo ErrCrystal
    Next
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    'Seteo la Query del reporte-----------------------------------------------------------------
    Cons = "SELECT Documento.DocCodigo , Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal, Documento.DocIVA, Documento.DocVendedor" _
            & " From " & paBD & ".dbo.Documento Documento " _
            & " Where DocCodigo = " & idDocumento
    If crSeteoSqlQuery(job_Moras%, Cons) = 0 Then GoTo ErrCrystal
        
    If Not TOPrinter Then If crMandoAPantalla(job_Moras, "Nota de Debito") = 0 Then GoTo ErrCrystal
    If TOPrinter Then If crMandoAImpresora(job_Moras, 1) = 0 Then GoTo ErrCrystal
    
    If Not crInicioImpresion(job_Moras, True, False) Then GoTo ErrCrystal
    
    If Not TOPrinter Then crEsperoCierreReportePantalla
    
    
    crCierroTrabajo job_Moras
    Screen.MousePointer = 0
    Exit Function

ErrCrystal:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError crMsgErr
    On Error Resume Next
    Screen.MousePointer = 11
    crCierroTrabajo job_Moras
    Screen.MousePointer = 0
End Function

Function BuscoDigitoUsuario(Codigo As Long) As String
On Error GoTo ErrBU
Dim Rs As rdoResultset

    BuscoDigitoUsuario = ""

    Cons = "SELECT * FROM USUARIO WHERE UsuCodigo = " & Codigo
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not Rs.EOF Then BuscoDigitoUsuario = Trim(Rs!UsuDigito)
    Rs.Close
    Exit Function
    
ErrBU:
End Function

Function fnc_DatosMoneda(Codigo As Integer, Optional retSigno As String, Optional ret_Nombre As String) As Boolean

    On Error GoTo ErrBU
    Dim Rs As rdoResultset
    fnc_DatosMoneda = False

    Cons = "SELECT * FROM Moneda WHERE MonCodigo = " & Codigo
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not Rs.EOF Then
        retSigno = Trim(Rs!monsigno)
        ret_Nombre = Trim(Rs!monnombre)
        fnc_DatosMoneda = True
    End If
    Rs.Close
    Exit Function
    
ErrBU:
End Function

Private Function EmitirCFE(ByVal Documento As clsDocumentoCGSA, ByVal CAE As clsCAEDocumento) As String
On Error GoTo errEC
    With New clsCGSAEFactura
        .URLAFirmar = prmURLFirmaEFactura
        .TasaBasica = TasaBasica
        .TasaMinima = TasaMinima
        .ImporteConInfoDeCliente = prmImporteConInfoCliente
        Set .Connect = cBase
        If Not .GenerarEComprobante(CAE, Documento, EmpresaEmisora, paCodigoDeSucursal) Then
            EmitirCFE = .XMLRespuesta
        End If
    End With
    Exit Function
errEC:
    EmitirCFE = "Error en firma: " & Err.Description
End Function

Private Function ValidarVersionEFactura() As Boolean
On Error GoTo errEC
    With New clsCGSAEFactura
        ValidarVersionEFactura = .ValidarVersion()
    End With
    Exit Function
errEC:
End Function

Private Sub CargoArticuloInteresesPorMora()
Dim rsArtMora As rdoResultset
        
    Set ProdInteresMora = New clsProducto
    Set rsArtMora = cBase.OpenResultset("SELECT ArtCodigo, ArtID, ArtTipo, ArtNombre, IvaCodigo, IvaDescripcion, IvaPorcentaje  FROM Articulo " _
                & "INNER JOIN ArticuloFacturacion ON ArtID = AFaArticulo " _
                & "INNER JOIN TipoIVA ON AFaIva = IvaCodigo " _
                & "WHERE ArtID = " & 557, rdOpenDynamic, rdConcurValues)
                
    With ProdInteresMora
        .ID = 5547
        .Nombre = Trim(rsArtMora("ArtNombre"))
        .TipoArticulo = rsArtMora("ArtTipo")
        .TipoIVA.Porcentaje = rsArtMora("IvaPorcentaje")
    End With
    rsArtMora.Close
        
End Sub


