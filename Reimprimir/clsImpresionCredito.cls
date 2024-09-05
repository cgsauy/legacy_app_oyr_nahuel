VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImpresionCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mDondeImprimo As clsParametrosImpresora
Public PathReportes As String
Public StringConnect As String

Public field_NombreDocumento As String
Public field_NombreRecibo As String
Public field_MonedaSimbolo As String
Public field_MonedaNombre As String
Public field_Hora As String
Public field_ClienteCedula As String
Public field_ClienteNombre As String
Public field_ClienteDireccion As String
Public field_ClienteGarantia As String
Public field_RUT As String
Public field_CreditoProxVto As String
Public field_CreditoFinanciacion As String
Public field_ReciboInfoCuota As String
Public field_ReciboInfoTFlete As String
Public field_ReciboImporteFlete As Currency
Public field_ReciboInfoVto As String
Public field_ReciboInfoSaldo As String
Public field_TextoRetira As String

Public Property Set DondeImprimo(ByVal value As clsParametrosImpresora)
    Set mDondeImprimo = value
End Property

Public Property Get DondeImprimo() As clsParametrosImpresora
    Set DondeImprimo = mDondeImprimo
End Property

Public Sub ImprimoFacturaContado_VSReport(ByVal IDDocumento As Long)
On Error GoTo errIS
Dim sQuery As String

    sQuery = "SELECT Top 1 '" & field_ClienteCedula & "' Cedula, '" & field_ClienteNombre & "' Cliente, '" & _
        field_ClienteDireccion & "' Direccion, '" & field_ClienteGarantia & "' Garantia, '" & field_RUT & "' RUC, " & IIf(field_RUT <> "", "''", "'X'") & " CFinal, '" & _
        field_NombreDocumento & "' NombreDocumento, '" & field_NombreRecibo & "' NombreRecibo, '" & field_MonedaNombre & "' MonedaNombre, '" & _
        field_MonedaSimbolo & "' MonedaSimbolo, Documento.DocCodigo, Documento.DocFecha, " & _
        "RTRIM(Documento.DocSerie) + '-' + RTRIM(Convert(varchar(7), Documento.DocNumero)) NumeroDocumento, " & _
        "Documento.DocTotal, '" & field_CreditoProxVto & "' ProxVtoCredito, '" & field_CreditoFinanciacion & "' Financiacion, '" & _
        field_TextoRetira & "' TextoRetira, dbo.FormatNumber(Documento.DocTotal - Documento.DocIVA, 2) NetoDocumento, " & _
        "dbo.FormatNumber(Documento.DocIVA, 2) IvaDocumento, dbo.FormatNumber(Documento.DocTotal, 2) TotalDocumento, " & _
        "'*'+RTRIM(CONVERT(varchar(2), Documento.DocTipo)) + 'D' + RTRIM(CONVERT(varchar(9), Documento.DocCodigo))+'*' CodigoBarra," & _
        "RTRIM(CONVERT(varchar(2), Documento.DocTipo)) + 'D' + RTRIM(CONVERT(varchar(9), Documento.DocCodigo)) InfoCodigoBarra," & _
        "Recibo.DocFecha RecFecha, CASE WHEN Recibo.DocSerie IS NULL THEN '' ELSE RTRIM(Recibo.DocSerie) + '-' + Convert(varchar(7), Recibo.DocNumero) END NumeroRecibo, " & _
        "Recibo.DocTotal RecTotal, '" & field_ReciboInfoCuota & "' InformacionCuota, '" & field_ReciboInfoTFlete & "' InformacionTFlete, '" & _
        field_ReciboInfoVto & "' InformacionVencimiento, '" & field_ReciboInfoSaldo & "' InformacionSaldo, " & _
        "dbo.FormatNumber(Recibo.DocTotal - " & field_ReciboImporteFlete & ", 2) ImporteCuota, " & _
        "dbo.FormatNumber(" & field_ReciboImporteFlete & ", 2) ImporteFlete, " & _
        "FROM  CGSA.dbo.Documento Documento LEFT OUTER JOIN CGSA.dbo.DocumentoPago DocumentoPago ON  Documento.DocCodigo = DocumentoPago.DPaDocASaldar " & _
        "LEFT OUTER JOIN CGSA.dbo.Documento Recibo ON  DocumentoPago.DPaDocQSalda = Recibo.DocCodigo WHERE Documento.DocCodigo = " & IDDocumento

    Dim frmPrint As New frmPrintReport
    Load frmPrint
    frmPrint.Visible = False
    
    With frmPrint.vsrReport
        .Clear                  ' clear any existing fields
        .FontName = "Tahoma"    ' set default font for all controls
        .FontSize = 8
        
        .Load PathReportes & "credito.xml", "Factura"
        
        .DataSource.ConnectionString = StringConnect
        .DataSource.RecordSource = sQuery
        
        Dim sQryRen As String
        sQryRen = "SELECT RTRIM(Convert(varchar(12), ArtCodigo)) + '  ' + RTRIM(ISNULL(AEsNombre, ArtNombre)) NombreArt, RenCantidad, dbo.FormatNumber(RenPrecio, 2) RenPrecio, dbo.FormatNumber(RenCantidad * RenPrecio, 2) TotalRenglon " & _
                " FROM Renglon INNER JOIN Articulo ON RenArticulo = ArtID" & _
                " LEFT OUTER JOIN ArticuloEspecifico On AEsTipoDocumento = 1 And AEsDocumento = RenDocumento And AEsArticulo = RenArticulo" & _
                " WHERE RenDocumento = " & IDDocumento
         
        .Fields("ReporteRenglon").Subreport.DataSource.ConnectionString = StringConnect
        .Fields("ReporteRenglon").Subreport.DataSource.RecordSource = sQryRen
        
        .Fields("ReporteRenglon2").Subreport.DataSource.ConnectionString = StringConnect
        .Fields("ReporteRenglon2").Subreport.DataSource.RecordSource = sQryRen
        
        frmPrint.vspPrinter.Device = DondeImprimo.Impresora
        frmPrint.vspPrinter.PaperBin = DondeImprimo.Bandeja
        frmPrint.vspPrinter.paperSize = DondeImprimo.Papel
        
        .Render frmPrint.vspPrinter
        
    End With
    frmPrint.vspPrinter.PrintDoc False
    
    Unload frmPrint
Exit Sub
errIS:
End Sub

Private Sub Class_Initialize()
On Error GoTo errCI

    field_Hora = Time
    field_ClienteNombre = ""
    field_ClienteDireccion = ""
    field_RUT = ""
    field_CFinal = "X"
    field_ClienteInferior = ""
    field_TextoRetira = ""
    
    Set DondeImprimo = New clsParametrosImpresora
    Exit Sub
    
errCI:
    MsgBox "Error al inicializar la clase: " & Err.Description, vbCritical, "Impresión Contado"
End Sub






