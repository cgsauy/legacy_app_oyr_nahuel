VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmResguardo 
   Caption         =   "Resguardos"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmResguardo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar tobMenu 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   1164
      ButtonWidth     =   1693
      ButtonHeight    =   1111
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "A emitir"
            Key             =   "find"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Emitir"
            Key             =   "save"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "A imprimir"
            Key             =   "findprint"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "print"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " "
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "A entregar"
            Key             =   "findsend"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Entregar"
            Key             =   "send"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsgLista 
      Height          =   1575
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   2778
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
      Left            =   0
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResguardo.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResguardo.frx":0B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResguardo.frx":0EAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResguardo.frx":1AFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResguardo.frx":1E7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResguardo.frx":21FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuGrilla 
      Caption         =   "MnuGrilla"
      Visible         =   0   'False
      Begin VB.Menu mnuGriMarcarTodos 
         Caption         =   "Marcar todos"
      End
      Begin VB.Menu mnuGriDesmarcarTodos 
         Caption         =   "Desmarcar todos"
      End
   End
End
Attribute VB_Name = "frmResguardo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DeshabilitoMenu()
    tobMenu.Buttons("send").Enabled = False
    tobMenu.Buttons("save").Enabled = False
    tobMenu.Buttons("print").Enabled = False
    tobMenu.Buttons("send").Enabled = False
End Sub

Private Function fnc_PrintDocumento(ByVal iDoc As Long) As Boolean
On Error GoTo errMPD
        
    Dim oPrint As New clsPrintManager
    With oPrint
        .SetDevice paIContadoN, paIContadoB, paPrintCtdoPaperSize
        If .LoadFileData(prmPathListados & "rptResguardos.txt") Then
            'fnc_PrintDocumento = .PrintDocumento(cBase, "EXEC prg_Resguardos 5, '" & iDoc & "', " & paCodigoDeUsuario)
            fnc_PrintDocumento = .PreviewPrint(cBase, "EXEC prg_Resguardos 5, '" & iDoc & "', " & paCodigoDeUsuario & ", 0")
            cBase.Execute "EXEC prg_Resguardos 6, '" & iDoc & "', " & paCodigoDeUsuario & ", 0"
        End If
    End With
    Set oPrint = Nothing
    Exit Function
errMPD:
    clsGeneral.OcurrioError "Error al imprimir el documento de código: " & iDoc, Err.Description, "Impresión de documentos"
End Function

Private Sub BuscarResguardosAImprirmirEntregar(ByVal accion As Byte)
On Error GoTo errBCP
    vsgLista.Rows = 1
    
    DeshabilitoMenu
    
    Dim rsR As rdoResultset
    Set rsR = cBase.OpenResultset("EXEC prg_Resguardos " & accion & ", null, " & paCodigoDeUsuario & ", 0", rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not rsR.EOF
        With vsgLista
            .AddItem ""
            .Cell(flexcpData, .Rows - 1, 0) = CStr(rsR("ComID"))
            If Not IsNull(rsR("EntGuardarComo")) Then
                .Cell(flexcpText, .Rows - 1, 1) = Trim(rsR("EntGuardarComo"))
                .Cell(flexcpData, .Rows - 1, 1) = CStr(rsR("ComProveedor"))
            End If
            If Not IsNull(rsR("ComFecha")) Then .Cell(flexcpText, .Rows - 1, 2) = Format(rsR("ComFecha"), "dd/mm/yyyy")
            If Not IsNull(rsR("ComNumero")) Then .Cell(flexcpText, .Rows - 1, 3) = Trim(rsR("ComNumero"))
            If Not IsNull(rsR("ComTotal")) Then
                .Cell(flexcpText, .Rows - 1, 4) = Trim(rsR("MonSimbolo")) + " " + Format(rsR("ComTotal"), "#,##0.00")
            End If
            .Cell(flexcpText, .Rows - 1, 5) = Format(rsR("ComTotal"), "#,##0.00")
        End With
        rsR.MoveNext
    Loop
    rsR.Close
    
    tobMenu.Buttons("print").Enabled = (vsgLista.Rows > 1 And accion = 3)
    tobMenu.Buttons("send").Enabled = (vsgLista.Rows > 1 And accion = 7)
    
    If vsgLista.Rows > 1 Then
        Call mnuGriMarcarTodos_Click
    End If
    
    Exit Sub

errBCP:
clsGeneral.OcurrioError "Error al buscar los comprobantes pendientes.", Err.Description, "Resguardo"
Screen.MousePointer = 0

End Sub

Private Sub ImprimirResguardoSeleccionados()
On Error GoTo errSub
Dim hayseleccion As Boolean
Dim fila As Integer

    hayseleccion = False
    
    For fila = 1 To vsgLista.Rows - 1
        If vsgLista.Cell(flexcpChecked, fila, 0) = flexChecked Then
            hayseleccion = True: Exit For
        End If
    Next
    
    If Not hayseleccion Then
        MsgBox "Debe seleccionar que resguardos  desea imprimir.", vbExclamation, "Imprimir resgurados"
        Exit Sub
    End If
    
     If MsgBox("¿Confirma imprimir los resguardos seleccionados?" & vbCrLf & vbCrLf & "Impresora: " & paIContadoN & vbCrLf & "Bandeja: " & paIContadoB & vbCrLf & "Papel: " & paPrintCtdoPaperSize _
        , vbQuestion + vbYesNo, "Resguardos") = vbNo Then
            Exit Sub
    End If
    
    Dim sprinterdefecto As String
    sprinterdefecto = Printer.DeviceName
    SeteoImpresoraPorDefecto paIContadoN
    
    For fila = 1 To vsgLista.Rows - 1
        If vsgLista.Cell(flexcpChecked, fila, 0) = flexChecked Then
            fnc_PrintDocumento Val(vsgLista.Cell(flexcpData, fila, 0))
        End If
    Next
    
    SeteoImpresoraPorDefecto sprinterdefecto
    vsgLista.Rows = 1
    Exit Sub
    
errSub:
    clsGeneral.OcurrioError "Error al intentar almacenar la información.", Err.Description, "Resguardos"
    Screen.MousePointer = 0
    Exit Sub
    
End Sub

Private Sub RealizarResguardos(ByVal accion As Byte)
On Error GoTo errSub
Dim documentos As String
Dim fila As Integer
    
    tobMenu.Buttons("print").Enabled = False
    
    Dim bMesAnterior As Boolean, bTomarCondicion As Boolean

    bTomarCondicion = (accion = 2)  'Quedamos en que si todos los comprobantes son del mes pasado se toma la acción x eso esta variable.
    
    bMesAnterior = False

    For fila = 1 To vsgLista.Rows - 1
        If vsgLista.Cell(flexcpChecked, fila, 0) = flexChecked Then
            documentos = documentos & IIf(documentos = "", "", ", ") & vsgLista.Cell(flexcpData, fila, 0)
            
            If accion = 2 Then
                If CDate(vsgLista.Cell(flexcpText, fila, 2)) < DateSerial(Year(Date), Month(Date), 1) Then
                    bMesAnterior = (True And bTomarCondicion)
                Else
                    bTomarCondicion = False
                    bMesAnterior = False
                End If
            End If
            
        End If
        
    Next
    
    If documentos = "" Then
        MsgBox "No hay comprobantes seleccionados para generar resguardos.", vbExclamation, "Resguardos"
        Exit Sub
    End If
    
    Dim result As VbMsgBoxResult
    result = vbNo
    
    Dim pregunta As String
    If accion = 4 Then
        If MsgBox("¿Confirma entregar los resguardos seleccionados?", vbQuestion + vbYesNo, "Resguardos") = vbNo Then
            Exit Sub
        End If
    Else
        If bMesAnterior And accion = 2 Then
            result = MsgBox("Los comprobantes son del mes anterior" & vbCrLf & vbCrLf & _
                        "Presione 'Si' para emitir el resguardo con fecha del último día del mes anterior" & vbCrLf & _
                        "Presione 'No' para emitir con el mes actual" & vbCrLf & _
                        "Presione 'Cancelar' para cancelar la emisión.", vbQuestion + vbYesNoCancel, "Emisión")
            If result = vbCancel Then Exit Sub
        Else
            If MsgBox("¿Confirma realizar los resguardos para los comprobantes seleccionados?", vbQuestion + vbYesNo, "Resguardos") = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    Screen.MousePointer = 11
    'INICIO TRANSACCION
    On Error GoTo errBT
    cBase.BeginTrans
    On Error GoTo errRB
    'cBase.Execute "EXEC prg_Resguardos  " & accion & ", '" & documentos & "', " & paCodigoDeUsuario & ", " & IIf(result = vbYes, 1, 0)
    'Todo esto es nuevo para la efactura antes estaba sólo la línea de arriba.
    If accion = 2 Then
        Dim rsD As rdoResultset
        Set rsD = cBase.OpenResultset("EXEC prg_Resguardos  " & accion & ", '" & documentos & "', " & paCodigoDeUsuario & ", " & IIf(result = vbYes, 1, 0), rdOpenDynamic, rdConcurValues)
        Do While Not rsD.EOF
            'Emito el CFE.
            EmitirCFE rsD(0)
            rsD.MoveNext
        Loop
        rsD.Close
    Else
        cBase.Execute "EXEC prg_Resguardos  " & accion & ", '" & documentos & "', " & paCodigoDeUsuario & ", " & IIf(result = vbYes, 1, 0)
    End If
    cBase.CommitTrans
    MsgBox "Se almaceno la información.", vbInformation, "Resguardos"
    If accion = 2 Then
        BuscarResguardosAImprirmirEntregar 3
    Else
        BuscarResguardosAImprirmirEntregar 7
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errSub:
    clsGeneral.OcurrioError "Error al intentar almacenar la información.", Err.Description, "Resguardos"
    Screen.MousePointer = 0
    Exit Sub
errBT:
    clsGeneral.OcurrioError "Error al iniciar la transacción", Err.Description, "Resguardos"
    Screen.MousePointer = 0
    Exit Sub

errT:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Error al grabar la información", Err.Description, "Resguardos"
    Screen.MousePointer = 0
    Exit Sub
    
errRB:
    Resume errT
    
End Sub

Private Sub BuscarComprobantesPendientes()
On Error GoTo errBCP
    vsgLista.Rows = 1
    tobMenu.Buttons("save").Enabled = False
    Dim rsR As rdoResultset
    Set rsR = cBase.OpenResultset("EXEC prg_Resguardos 1, null, " & paCodigoDeUsuario & ", 0", rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not rsR.EOF
        With vsgLista
            .AddItem ""
            .Cell(flexcpData, .Rows - 1, 0) = CStr(rsR("ComID"))
            If Not IsNull(rsR("EntGuardarComo")) Then
                .Cell(flexcpText, .Rows - 1, 1) = Trim(rsR("EntGuardarComo"))
                .Cell(flexcpData, .Rows - 1, 1) = CStr(rsR("ComProveedor"))
            End If
            If Not IsNull(rsR("ComFecha")) Then .Cell(flexcpText, .Rows - 1, 2) = Format(rsR("ComFecha"), "dd/mm/yyyy")
            If Not IsNull(rsR("ComNumero")) Then .Cell(flexcpText, .Rows - 1, 3) = Trim(rsR("ComNumero"))
            If Not IsNull(rsR("ComTotal")) Then
                .Cell(flexcpText, .Rows - 1, 4) = Trim(rsR("MonSimbolo")) + " " + Format(rsR("ComTotal"), "#,##0.00")
            End If
            .Cell(flexcpText, .Rows - 1, 5) = Format(rsR("IVA"), "#,##0.00")
        End With
        rsR.MoveNext
    Loop
    rsR.Close
    tobMenu.Buttons("save").Enabled = (vsgLista.Rows > 1)
    Exit Sub
errBCP:
clsGeneral.OcurrioError "Error al buscar los comprobantes pendientes.", Err.Description, "Resguardo"
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    ObtengoSeteoForm Me, Me.Left, Me.Top, Me.Width, Me.Height
    With vsgLista
        .Cols = 1
        .Rows = 1
        .FormatString = "|Proveedor|^Fecha|Factura|>Importe|>IVA 50%|"
        .FixedCols = 0
        .ColWidth(0) = 400
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        .ColWidth(4) = 1600
        .ColWidth(5) = 1200
        .ColWidth(6) = 0
        .BackColorBkg = vbWindowBackground
        .SheetBorder = vbWindowBackground
        .ColDataType(0) = flexDTBoolean
        .Editable = True
    End With
    DeshabilitoMenu
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
On Error Resume Next
    vsgLista.Move 0, tobMenu.Height, ScaleWidth, ScaleHeight - tobMenu.Height
    'Ajusto las columnas.
    vsgLista.ColWidth(1) = vsgLista.Width - (vsgLista.ColWidth(0) + vsgLista.ColWidth(2) + vsgLista.ColWidth(3) + vsgLista.ColWidth(4) + vsgLista.ColWidth(5) + 360)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    cBase.Close
    GuardoSeteoForm Me
End Sub

Private Sub mnuGriDesmarcarTodos_Click()
Dim fila As Integer
    For fila = 1 To vsgLista.Rows - 1
        vsgLista.Cell(flexcpChecked, fila, 0) = flexUnchecked
    Next
End Sub

Private Sub mnuGriMarcarTodos_Click()
Dim fila As Integer
    For fila = 1 To vsgLista.Rows - 1
        vsgLista.Cell(flexcpChecked, fila, 0) = flexChecked
    Next
End Sub

Private Sub tobMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case LCase(Button.Key)
        Case "find": BuscarComprobantesPendientes
        Case "save": RealizarResguardos 2
        Case "send": RealizarResguardos 4
        Case "findprint": BuscarResguardosAImprirmirEntregar 3
        Case "findsend": BuscarResguardosAImprirmirEntregar 7
        Case "print": ImprimirResguardoSeleccionados
    End Select
End Sub

Private Sub vsgLista_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col <> 0)
End Sub

Private Sub vsgLista_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If vsgLista.Row > 0 And KeyCode = vbKeySpace Then
        If vsgLista.Cell(flexcpChecked, vsgLista.Row, 0) = flexUnchecked Then
            vsgLista.Cell(flexcpChecked, vsgLista.Row, 0) = flexChecked
        Else
            vsgLista.Cell(flexcpChecked, vsgLista.Row, 0) = flexUnchecked
        End If
    ElseIf vsgLista.Rows > 1 And KeyCode = 93 Then
        PopupMenu mnuGrilla
    End If
End Sub

Private Sub vsgLista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And vsgLista.Rows > 1 Then
        PopupMenu mnuGrilla
    End If
End Sub

Private Sub vsgLista_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col <> 0)
End Sub

Private Sub SeteoImpresoraPorDefecto(DeviceName As String)
Dim X As Printer

    For Each X In Printers
        If Trim(X.DeviceName) = Trim(DeviceName) Then
            Set Printer = X
            Exit For
        End If
    Next
    
End Sub

Private Function EmitirCFE(ByVal idDocumento As Long) As String
On Error GoTo errEC
    With New clsCGSAEFactura
        Set .Connect = cBase
        If Not .GenerarEComprobanteZureo(idDocumento) Then
            EmitirCFE = .XMLRespuesta
        End If
    End With
    Exit Function
errEC:
    EmitirCFE = "Error al firmar: " & Err.Description
End Function
