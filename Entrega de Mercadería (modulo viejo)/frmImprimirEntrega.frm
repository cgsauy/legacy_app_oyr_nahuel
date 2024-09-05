VERSION 5.00
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Begin VB.Form frmImprimirEntrega 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir Entregas"
   ClientHeight    =   6240
   ClientLeft      =   2685
   ClientTop       =   2175
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImprimirEntrega.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   7680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton bImprimir 
      Caption         =   "&Imprimir"
      Height          =   315
      Left            =   5340
      TabIndex        =   5
      Top             =   300
      Width           =   975
   End
   Begin VB.ComboBox cEntregas 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2700
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   300
      Width           =   2475
   End
   Begin VB.TextBox tCBarra 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   300
      Width           =   2415
   End
   Begin vsViewLib.vsPrinter vsPrint 
      Height          =   5235
      Left            =   60
      TabIndex        =   6
      Top             =   900
      Width           =   7575
      _Version        =   196608
      _ExtentX        =   13361
      _ExtentY        =   9234
      _StockProps     =   229
      Appearance      =   1
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ConvInfo        =   1418783674
      Zoom            =   50
   End
   Begin VB.Label lDocumento 
      BackStyle       =   0  'Transparent
      Caption         =   "&Factura o Remito"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   660
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha/Hora Entrega"
      Height          =   255
      Left            =   2700
      TabIndex        =   3
      Top             =   60
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Factura o Remito"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1575
   End
End
Attribute VB_Name = "frmImprimirEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mSQL As String
Dim rsXX As rdoResultset

Dim prmTipoDoc As Integer, prmIDDocumento As Long
Dim prmNombreCliente As String

'----------------------------------------------------------------------------------
'   Interpreta el Texto del Codigo de Barras
'   Formato:    XDXXXX          TipoDocumento   D Numero de Documento
'----------------------------------------------------------------------------------
Private Sub FormatoBarras(Texto As String)

Dim aCodDoc As Long
    
    On Error GoTo errInt
    Texto = UCase(Texto)
    
    '1) Veo si es x codigo de barras o x ids de documento
    If (Mid(Texto, 2, 1) = "D" And IsNumeric(Mid(Texto, 1, 1)) And Len(Texto) > 3) Or (Mid(Texto, 3, 1) = "D" And IsNumeric(Mid(Texto, 1, 2)) And Len(Texto) > 4) Then   'Codigo de Barras
        prmTipoDoc = CLng(Mid(Texto, 1, InStr(Texto, "D") - 1))
        aCodDoc = CLng(Trim(Mid(Texto, InStr(Texto, "D") + 1, Len(Texto))))
    Else
        'Puso Serie y Numero de Documento o Numero de Remito
        If IsNumeric(Texto) Then
            prmTipoDoc = TipoDocumento.Remito        'Remito
            aCodDoc = Texto
        Else
            'Puso Serie y Numero de Documento
             If Not zfn_BuscoDocPorTexto(Texto, aCodDoc, prmTipoDoc) Then aCodDoc = -1
        End If
        
    End If
    
    If aCodDoc = -1 Then
        MsgBox "No existe un documento que coincida con los valores ingresados.", vbExclamation, "No hay Datos"
        Exit Sub
    End If
    
    Select Case prmTipoDoc
        Case TipoDocumento.Remito: fnc_BuscoRemito aCodDoc
        
        Case TipoDocumento.Contado, TipoDocumento.Credito: fnc_BuscoDocumento aCodDoc
        
        Case TipoDocumento.NotaCredito, TipoDocumento.NotaDevolucion, TipoDocumento.NotaEspecial
                        fnc_BuscoDocumento aCodDoc
                        
        Case Else
            MsgBox "El código de barras ingresado no es correcto. El documento no coincide con los predefinidos.", vbCritical, "ATENCIÓN"
            prmIDDocumento = 0
    End Select
    
    cEntregas.Clear
    If prmIDDocumento > 0 Then fnc_CargoEntregas
    
    Screen.MousePointer = 0
    Exit Sub
    
errInt:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al interpretar el código de barras.", Err.Description
End Sub

Private Function fnc_BuscoDocumento(idDocumento As Long)

Dim idCliente As Long

    Screen.MousePointer = 11
    idCliente = 0
    
    lDocumento.Caption = ""
    
    Cons = "Select * from Documento " & _
               " Where DocCodigo = " & idDocumento
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not RsAux.EOF Then
        prmIDDocumento = RsAux!DocCodigo
        
        lDocumento.Caption = UCase(NombreDocumento(RsAux!DocTipo) & " " & Trim(RsAux!DocSerie) & "-" & RsAux!DocNumero)
                
        idCliente = RsAux!DocCliente
       
        If RsAux!DocAnulado Then
            Screen.MousePointer = 0
            MsgBox "El documento ingresado ha sido anulado. Verifique", vbCritical, "DOCUMENTO ANULADO"
        Else
            If Not IsNull(RsAux!DocPendiente) Then
                Screen.MousePointer = 0
                MsgBox "La mercadería está pendiente de entrega. Verifique", vbInformation, "ATENCIÓN"
            End If
        End If
        
    Else
        Screen.MousePointer = 0
        prmIDDocumento = 0
        MsgBox "No existe un documento para las características ingresadas.", vbExclamation, "ATENCIÓN"
    End If
    RsAux.Close
        
    If prmIDDocumento <> 0 Then prmNombreCliente = fnc_CargoCliente(idCliente)
            
End Function

Private Sub fnc_BuscoRemito(Numero As Long)

Dim idCliente As Long

    Cons = "Select * from Remito, Documento" _
            & " Where RemCodigo = " & Numero _
            & " And RemDocumento = DocCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)

    If Not RsAux.EOF Then
        'prmIDDocumento = RsAux!DocCodigo
        prmIDDocumento = RsAux!RemCodigo
        lDocumento.Caption = "REMITO Nº " & RsAux!RemCodigo
        idCliente = RsAux!DocCliente
        
        '--------------------------------------------------------------------------------------------------------------------------------------
        If RsAux!DocAnulado Then
            Screen.MousePointer = 0
            MsgBox "El documento ingresado ha sido anulado. Verifique", vbCritical, "DOCUMENTO ANULADO"
        Else
            If Not IsNull(RsAux!DocPendiente) Then
                Screen.MousePointer = 0
                MsgBox "La mercadería está pendiente de entrega. Verifique", vbInformation, "ATENCIÓN"
            End If
        End If
        '--------------------------------------------------------------------------------------------------------------------------------------
        
    Else
        Screen.MousePointer = 0
        prmIDDocumento = 0
        MsgBox "No existe un remito para las características ingresadas.", vbExclamation, "ATENCIÓN"
    End If
    
    RsAux.Close

    prmNombreCliente = fnc_CargoCliente(idCliente)
    
End Sub

Private Function fnc_CargoCliente(Cliente As Long) As String
    
    On Error GoTo errCliente
    fnc_CargoCliente = ""
    
    Cons = "Select CliCiRuc, CliTipo, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2) " _
           & " From Cliente, CPersona " _
           & " Where CliCodigo = " & Cliente _
           & " And CliCodigo = CPeCliente " _
                                                & " UNION " _
           & " Select CliCiRuc, CliTipo, Nombre = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')')  " _
           & " From Cliente, CEmpresa " _
           & " Where CliCodigo = " & Cliente _
           & " And CliCodigo = CEmCliente"

    Set rsXX = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsXX.EOF Then fnc_CargoCliente = Trim(rsXX!Nombre)
    rsXX.Close
    
    Exit Function
    
errCliente:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al buscar los datos del cliente."
End Function

Private Sub bImprimir_Click()
    
    If prmIDDocumento = 0 Then tCBarra.SetFocus: Exit Sub
    If cEntregas.ListIndex = -1 Then Exit Sub
     
    fnc_Imprimir
    
End Sub

Private Sub cEntregas_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cEntregas.ListIndex <> -1 Then bImprimir.SetFocus
    End If
    
End Sub

Private Sub Form_Load()
On Error Resume Next
    tCBarra.Text = "": lDocumento.Caption = "": cEntregas.Clear
    
    With vsPrint
        .PhysicalPage = True
        .PaperSize = 1
        .Orientation = orPortrait
        .MarginTop = 550: .MarginRight = 550: .MarginLeft = 550
        .PageBorder = pbNone
        .AbortWindow = False
    End With

End Sub

Private Function fnc_CargoEntregas()
Dim sHrs As String
    cEntregas.Clear
    
    'MSFFecha , MSFTipoLocal, MSFLocal, MSFArticulo, MSFCantidad, MSFEstado, MSFTipoDocumento, MSFDocumento, MSFUsuario, MSFTerminal
    mSQL = "Select Distinct(MSFFecha) from MovimientoStockFisico" & _
                " Where MSFTipoDocumento = " & prmTipoDoc & " And MSFDocumento = " & prmIDDocumento & _
                " And MSFTipoLocal =" & TipoLocal.Deposito & " And MSFLocal =" & paCodigoDeSucursal & _
                " Order by MSFFecha DESC"
    
    Set rsXX = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    Do While Not rsXX.EOF
        If InStr(1, "," & sHrs & ",", "," & Format(rsXX(0), "dd/mm/yyyy HH:nn:ss") & ",", vbTextCompare) = 0 Then
            cEntregas.AddItem Format(rsXX(0), "dd/mm/yyyy HH:nn:ss")
            sHrs = sHrs & IIf(sHrs <> "", ",", "") & Format(rsXX(0), "dd/mm/yyyy HH:nn:ss")
        End If
        rsXX.MoveNext
    Loop
    rsXX.Close
    
    If cEntregas.ListCount > 0 Then cEntregas.ListIndex = 0

End Function

Private Function fnc_Imprimir()
On Error GoTo errImprimir
Dim mSQL As String
Dim rs1 As rdoResultset

Dim v_Articulos As String, v_Fecha As String, v_Hasta As String
    
    v_Fecha = Format(cEntregas.Text, "mm/dd/yyyy hh:mm:ss")
    v_Hasta = DateAdd("s", 1, CDate(cEntregas.Text))
    v_Hasta = Format(v_Hasta, "mm/dd/yyyy hh:mm:ss")
    
    mSQL = "Select MSFFecha, MSFCantidad, ArtCodigo, ArtNombre from MovimientoStockFisico, Articulo " & _
                " Where MSFArticulo = ArtID " & _
                " And MSFFecha Between '" & v_Fecha & "' AND '" & v_Hasta & "' " & _
                " And MSFTipoDocumento = " & prmTipoDoc & " And MSFDocumento = " & prmIDDocumento & _
                " And MSFTipoLocal =" & TipoLocal.Deposito & " And MSFLocal =" & paCodigoDeSucursal
    
    Set rs1 = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    If Not rs1.EOF Then
        v_Fecha = Format(rs1!MSFFecha, "dd/mm/yyyy hh:mm")
        Do While Not rs1.EOF
            
            v_Articulos = v_Articulos & IIf(v_Articulos = "", "", "|") & _
                               Abs(rs1!MSFCantidad) & "  " & Format(rs1!ArtCodigo, "(#,000,000)") & " " & Trim(rs1!ArtNombre)
                
            rs1.MoveNext
        Loop
    End If
    rs1.Close
        
    With vsPrint
        '.Visible = True: .ZOrder 0
        .PhysicalPage = False
        .FileName = "Impresion de Entrega"

        .StartDoc
        
        .TextAlign = taCenterBaseline
        .Font.Name = "Tahoma": .Font.Size = 16
        
        .Paragraph = "ENTREGA DE MERCADERIA"
        .Paragraph = " "
        
        .TextAlign = taLeftBaseline
        .Font.Size = 12: .Font.Bold = False
        .Paragraph = "Local de Entrega: " & prmNombreLocal
        .Paragraph = "Fecha / Hora: " & v_Fecha
        .Paragraph = "Factura o Remito: " & lDocumento.Caption
        .Paragraph = "Nombre: " & prmNombreCliente
        .MarginLeft = 1500
        
        .Paragraph = " ":
        .Font.Bold = True
        .Paragraph = UCase("Artículos Entregados")
        .Font.Bold = False
        If v_Articulos <> "" Then
            Dim arrTMP() As String, iX
            arrTMP = Split(v_Articulos, "|")
            For iX = LBound(arrTMP) To UBound(arrTMP)
                .Paragraph = arrTMP(iX)
            Next
        End If
        
        .MarginLeft = 550
        .Paragraph = " ": .Paragraph = " "
        .Paragraph = "Recibí conforme "
        .Paragraph = " ": .Paragraph = " "
        .Paragraph = "Nombre:  ____________________________________________"
        .Paragraph = " ": .Paragraph = " ": .Paragraph = " "
        .Paragraph = "Firma:   ____________________________________________"
        .Paragraph = " ": .Paragraph = " " ': .Paragraph = " "
        .Paragraph = "Cédula:   _______________________"
        
        .EndDoc
            
        .Copies = 1
        .PrintDoc True
        
'        tCBarra.Text = "": lDocumento.Caption = "": cEntregas.Clear
        
    End With
    Screen.MousePointer = 0
    
    Exit Function
errImprimir:
    clsGeneral.OcurrioError "Error al cargar los datos para imprimir.", Err.Description
    Screen.MousePointer = 0
End Function


Private Sub tCBarra_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tCBarra.Text) <> "" Then
            FormatoBarras Trim(tCBarra.Text)
            On Error Resume Next
            If cEntregas.ListCount > 0 Then cEntregas.SetFocus
        End If
    End If
    
End Sub

Private Function zfn_BuscoDocPorTexto(adTexto As String, retIDDoc As Long, retIDTipoD) As Boolean
On Error GoTo errDoc
    zfn_BuscoDocPorTexto = False
    
    Dim mDSerie As String, mDNumero As Long
    Dim adQ As Integer, adCodigo As Long, adTipoD As Integer
        
    If InStr(adTexto, "-") <> 0 Then
        mDSerie = Mid(adTexto, 1, InStr(adTexto, "-") - 1)
        mDNumero = Val(Mid(adTexto, InStr(adTexto, "-") + 1))
    Else
        mDSerie = Mid(adTexto, 1, 1)
        mDNumero = Val(Mid(adTexto, 2))
    End If
    
    adTexto = UCase(mDSerie) & "-" & mDNumero
        
    Screen.MousePointer = 11
    adQ = 0: adTexto = ""
    
    'Cargo combo con tipos de docuemento--------------------------------------
    Cons = "Select DocCodigo, DocTipo, DocFecha as Fecha, DocSerie as Serie, Convert(char(7),DocNumero) as Numero " & _
               " From Documento " & _
               " Where DocSerie = '" & mDSerie & "'" & _
               " And DocNumero = " & mDNumero & _
               " And DocTipo IN (" & TipoDocumento.Contado & ", " & TipoDocumento.Credito & ", " & TipoDocumento.NotaCredito & ", " & _
                                               TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial & ")"
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        adCodigo = RsAux!DocCodigo
        adTipoD = RsAux!DocTipo
        adQ = 1
        RsAux.MoveNext: If Not RsAux.EOF Then adQ = 2
    End If
    RsAux.Close
        
        Select Case adQ
            Case 2
                Dim miLDocs As New clsListadeAyuda
                If miLDocs.ActivarAyuda(cBase, Cons, 4100, 2) <> 0 Then
                    adCodigo = miLDocs.RetornoDatoSeleccionado(0)
                    adTipoD = miLDocs.RetornoDatoSeleccionado(1)
                End If
                Set miLDocs = Nothing
                Me.Refresh
        End Select
        
        If adCodigo > 0 Then
            zfn_BuscoDocPorTexto = True
            retIDDoc = adCodigo
            retIDTipoD = adTipoD
        Else
            zfn_BuscoDocPorTexto = False
        End If
        
        Screen.MousePointer = 0
    
    Exit Function
errDoc:
    clsGeneral.OcurrioError "Error al buscar el documento.", Err.Description
    Screen.MousePointer = 0
End Function


