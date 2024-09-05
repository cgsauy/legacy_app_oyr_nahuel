VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRePrint 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Reimpresión de remitos"
   ClientHeight    =   5550
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRePrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsDocs 
      Height          =   1335
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   2355
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
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRePrint.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRePrint.frx":065C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   1005
      ButtonWidth     =   1402
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "print"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Key             =   "exit"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código de impresión o de envío:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lbMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Marque en la grilla los documentos que desea reimprimir (con botón derecho marca y desmarca todos)."
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
      TabIndex        =   3
      Top             =   4560
      Width           =   7695
   End
   Begin VB.Shape shfac 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      FillColor       =   &H00DCFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   4440
      Width           =   8340
   End
   Begin VB.Menu MnuBotonD 
      Caption         =   "botond"
      Visible         =   0   'False
      Begin VB.Menu MnuBDMarcar 
         Caption         =   "Marcar todos"
      End
      Begin VB.Menu MnuBDDesmarcar 
         Caption         =   "Desmarcar todos"
      End
   End
End
Attribute VB_Name = "frmRePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function GetTipoDocumento(ByVal doc As Long) As Long
Dim rsd As rdoResultset
    Cons = "SELECT DocTipo FROM Documento WHERE DocCodigo = " & doc
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    GetTipoDocumento = RsAux(0)
    RsAux.Close
End Function

Private Sub ImprimirDocumentos()
Dim iRow As Integer
On Error GoTo errID
    If MsgBox("¿Confirma reimprimir los documentos seleccionados?", vbQuestion + vbYesNo, "Reimprimir") = vbYes Then
        
        Dim iUsrSuceso As Long
        Dim sDefSuceso As String
        
        On Error Resume Next
        Screen.MousePointer = 11
        Dim objSuceso As New clsSuceso
        objSuceso.ActivoFormulario paCodigoDeUsuario, "Reimpresión de Documentos", cBase
        iUsrSuceso = objSuceso.RetornoValor(Usuario:=True)
        sDefSuceso = objSuceso.RetornoValor(Defensa:=True)
        Set objSuceso = Nothing
        Me.Refresh
        If iUsrSuceso = 0 Then MsgBox "No se imprimirá ningún documento.", vbInformation, "Atención": Screen.MousePointer = 0: Exit Sub 'Abortó el ingreso del suceso
        '---------------------------------------------------------------------------------------------
        
        For iRow = 1 To vsDocs.Rows - 1
            If vsDocs.Cell(flexcpChecked, iRow, 0) = flexChecked Then
                        'TipoSuceso.Reimpresiones = 10
                objGral.RegistroSuceso cBase, Now, 10, paCodigoDeTerminal, iUsrSuceso, CLng(vsDocs.Cell(flexcpData, iRow, 0)), _
                                           Descripcion:=vsDocs.Cell(flexcpText, iRow, 1), Defensa:=Trim(sDefSuceso)
    
                If GetTipoDocumento(vsDocs.Cell(flexcpData, iRow, 0)) = 6 Then
                    'frmDistribuirEnvio.EsperaConEvento
                    frmDistribuirEnvio.fnc_PrintDocumento vsDocs.Cell(flexcpData, iRow, 0)
                Else
                    frmDistribuirEnvio.ImprimoEFactura vsDocs.Cell(flexcpData, iRow, 0), vsDocs.Cell(flexcpData, iRow, 2), (Val(vsDocs.Cell(flexcpData, iRow, 1)) <> paCodigoDeSucursal), Val(vsDocs.Cell(flexcpData, iRow, 3))
                End If
                
            End If
        Next
    End If
Exit Sub
errID:
    objGral.OcurrioError "Error al imprimir.", Err.Description
End Sub

Private Function f_GetDireccionRsAux() As String
    If Not IsNull(RsAux!CalNombre) Then
        f_GetDireccionRsAux = Trim(RsAux!CalNombre) & " " & Trim(RsAux!DirPuerta)
        If Not IsNull(RsAux!DirLetra) Then f_GetDireccionRsAux = f_GetDireccionRsAux & " " & Trim(RsAux!DirLetra)
        If Not IsNull(RsAux!DirApartamento) Then f_GetDireccionRsAux = f_GetDireccionRsAux & " / " & Trim(RsAux!DirApartamento)
    End If
End Function

Private Sub db_FillGridDatosCodImpresion()
On Error GoTo errGDR
Dim QTotal As Integer, QCamion As Integer
Dim lLastID As Long, lLastVC As Long
Dim sFM As String
    
    'Busco los datos de la tabla repartoimpresión.
    Screen.MousePointer = 11
    vsDocs.Rows = 1
    Toolbar1.Buttons("print").Enabled = False
        
    Dim iCodImpresion As Long
    'Verifico si es un envío o un código de impresión
    Cons = "SELECT EnvCodImpresion FROM Envio WHERE (EnvCodigo = " & txtCodigo.Text & " OR EnvCodImpresion = " & txtCodigo.Text _
        & ") And EnvCodImpresion Is NOT NULL"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        Screen.MousePointer = 0
        MsgBox "No existe un envío o un código de impresión para la información ingresada.", vbExclamation, "Atención"
        RsAux.Close
        Exit Sub
    Else
        iCodImpresion = RsAux(0)
    End If
    RsAux.Close
    
    Dim iCodDoc As Long     'Se puede repetir el #Doc por el vacon.
    cBase.QueryTimeout = 30

' Cons = "SELECT DocCodigo, DocTipo, DocSerie, DocNumero, DocTotal, EnvCodigo, CalNombre, DirPuerta, DirLetra, DirApartamento" & _
        " FROM Envio, Documento, Direccion, Calle " & _
        " WHERE EnvCodimpresion = " & iCodImpresion & _
        " AND (EnvDocumento = DocCodigo OR (EnvDocumentoFactura = DocCodigo AND EnvFormapago <> 1))" & _
        " OR EnvDocumento IN (SELECT RDoRemito From RemitoDocumento, VentaTelefonica WHERE VTeDocumento = RDoDocumento)" & _
        " AND EnvDireccion = DirCodigo And DirCalle = CalCodigo"
    
    '" Order by DocCodigo"
    Set RsAux = cBase.OpenResultset("prg_DistribuirEnvio_ReimprimirDocumentos " & iCodImpresion, rdOpenDynamic, rdConcurValues)
    
    
    Do While Not RsAux.EOF

        With vsDocs
            .AddItem ""
            .Cell(flexcpChecked, .Rows - 1, 0) = flexChecked
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux("Documento"))
            .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux("Domicilio"))
            If RsAux("Sucursal") <> paCodigoDeSucursal Then .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = &H8080FF
            
            .Cell(flexcpData, .Rows - 1, 0) = Trim(RsAux("DocID"))
            .Cell(flexcpData, .Rows - 1, 1) = Trim(RsAux("Sucursal"))
            .Cell(flexcpData, .Rows - 1, 2) = Trim(RsAux("Envio"))
            .Cell(flexcpData, .Rows - 1, 3) = Trim(CStr(RsAux("TipoDoc")))

        End With
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    Toolbar1.Buttons("print").Enabled = (vsDocs.Rows > vsDocs.FixedRows)
    Screen.MousePointer = 0
    Exit Sub
errGDR:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al buscar y cargar la información del código de impresión.", Err.Description
End Sub


Private Sub Form_Load()
    With vsDocs
        .Rows = 1
        .Cols = 1
        .FixedRows = 1
        .FormatString = "Imprimir|Documento|Dirección|"
        .FixedCols = 0
        .Editable = True
        .ColHidden(3) = True
        .ColDataType(0) = flexDTBoolean
        .ExtendLastCol = True
        .RowHeight(0) = 315
        .ColWidth(1) = 2000
        .BackColorSel = vbInfoBackground
        .ForeColorSel = vbWindowText
        .FocusRect = flexFocusHeavy
        .BorderStyle = flexBorderNone
        .SheetBorder = .BackColorBkg
    End With
    Toolbar1.Buttons("print").Enabled = False
    
'    frmDistribuirEnvio.ImprimoEFactura 14566624, 0, True, 48
'    frmDistribuirEnvio.ImprimoEFactura 14566625, 0, True, 47

    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Me.shfac.Move 15, Me.ScaleHeight - Me.shfac.Height, Me.ScaleWidth - 30
    Me.lbMsg.Move 120, shfac.Top + 120, shfac.Width - 240
    vsDocs.Move 30, vsDocs.Top, ScaleWidth - 60, shfac.Top - vsDocs.Top
End Sub

Private Sub MnuBDDesmarcar_Click()
    On Error Resume Next
    vsDocs.Cell(flexcpChecked, 1, 0, vsDocs.Rows - 1, 0) = flexNoCheckbox
End Sub

Private Sub MnuBDMarcar_Click()
On Error Resume Next
    vsDocs.Cell(flexcpChecked, 1, 0, vsDocs.Rows - 1, 0) = flexChecked
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "print"
            ImprimirDocumentos
        Case "exit"
            Unload Me
    End Select
End Sub

Private Sub txtCodigo_Change()
    If Val(txtCodigo.Tag) > 0 Then
        txtCodigo.Tag = 0
        vsDocs.Rows = 1
        Toolbar1.Buttons("print").Enabled = False
    End If
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
On Error GoTo errKP
    If KeyAscii = vbKeyReturn Then
        If Val(txtCodigo.Tag) = 0 Then
            db_FillGridDatosCodImpresion
        Else
            vsDocs.SetFocus
        End If
    End If
    Exit Sub
errKP:
    objGral.OcurrioError "Error al buscar los documentos.", Err.Description, "Buscar documentos"
End Sub

Private Sub vsDocs_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col <> 0)
End Sub

Private Sub vsDocs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vsDocs.Rows > 1 And Button = 2 Then
        PopupMenu MnuBotonD
    End If
End Sub
