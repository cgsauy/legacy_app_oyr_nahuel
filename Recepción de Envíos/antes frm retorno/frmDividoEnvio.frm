VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDividoEnvio 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dividir un envío"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5580
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
   ScaleHeight     =   4065
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   3600
      TabIndex        =   5
      Top             =   60
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   582
      ButtonWidth     =   1588
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
            Key             =   "save"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "exit"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.TextBox tCodigo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      MaxLength       =   8
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsArticulos 
      Height          =   2295
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4048
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
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
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
      Left            =   1560
      Top             =   3360
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
            Picture         =   "frmDividoEnvio.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDividoEnvio.frx":0112
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbDireccion 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Envío:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
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
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   5295
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
      Top             =   3240
      Width           =   5460
   End
End
Attribute VB_Name = "frmDividoEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public prmEnvio As Long

Private Function CopioDireccion(ByVal lnCodDireccion As Long) As Long
Dim aIdCalle As Long, aNroPuerta As Long

    'Copio la Direccion
    Dim RsDO As rdoResultset
    Dim RsDC As rdoResultset
    
    'Direccion ORIGINAL
    Cons = "Select * from Direccion Where DirCodigo = " & lnCodDireccion
    Set RsDO = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    'Direccion COPIA
    Cons = "Select * from Direccion Where DirCodigo = 0"
    Set RsDC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsDC.EOF Then
        RsDC.Edit
    Else
        RsDC.AddNew
    End If
    If Not IsNull(RsDO!DirComplejo) Then RsDC!DirComplejo = RsDO!DirComplejo
    
    RsDC!DirCalle = RsDO!DirCalle
    aIdCalle = RsDO!DirCalle
    
    RsDC!DirPuerta = RsDO!DirPuerta
    aNroPuerta = RsDO!DirPuerta
    
    RsDC!DirBis = RsDO!DirBis
    If Not IsNull(RsDO!DirLetra) Then RsDC!DirLetra = RsDO!DirLetra
    If Not IsNull(RsDO!DirApartamento) Then RsDC!DirApartamento = RsDO!DirApartamento
    
    If Not IsNull(RsDO!DirCampo1) Then RsDC!DirCampo1 = RsDO!DirCampo1
    If Not IsNull(RsDO!DirSenda) Then RsDC!DirSenda = RsDO!DirSenda
    If Not IsNull(RsDO!DirCampo2) Then RsDC!DirCampo2 = RsDO!DirCampo2
    If Not IsNull(RsDO!DirBloque) Then RsDC!DirBloque = RsDO!DirBloque
    
    If Not IsNull(RsDO!DirEntre1) Then RsDC!DirEntre1 = RsDO!DirEntre1
    If Not IsNull(RsDO!DirEntre2) Then RsDC!DirEntre2 = RsDO!DirEntre2
    If Not IsNull(RsDO!DirAmpliacion) Then RsDC!DirAmpliacion = RsDO!DirAmpliacion
    RsDC!DirConfirmada = RsDO!DirConfirmada
    If Not IsNull(RsDO!DirVive) Then RsDC!DirVive = RsDO!DirVive
    RsDC.Update
    RsDC.Close: RsDO.Close
    
    Cons = "Select Max(DirCodigo) from Direccion Where DirCalle = " & aIdCalle _
        & " And DirPuerta = " & aNroPuerta
    Set RsDC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    CopioDireccion = RsDC(0)
    RsDC.Close

End Function


Private Sub loc_FindEnvio()
On Error GoTo errFE
Dim lAux As Long
Dim sVaCon As String

    Screen.MousePointer = 11
    Toolbar1.Buttons("save").Enabled = False
    vsArticulos.Rows = 1
    Cons = "Select EnvCodigo, IsNull(EnvVaCon, 0) as VaCon, EnvEstado, EnvFModificacion, EnvDireccion" & _
                " From Envio " & _
                " Where EnvCodigo = " & Val(tCodigo.Text)
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        RsAux.Close
        Screen.MousePointer = 0
        MsgBox "No existe un envío con ese código.", vbExclamation, "Atención"
        Exit Sub
    Else
        If RsAux("EnvEstado") <> 3 Then
            Screen.MousePointer = 0
            RsAux.Close
            MsgBox "El envío no tiene el estado impreso, para modificarlo acceda al formulario de envíos.", vbExclamation, "Atención"
            Exit Sub
        Else
            tCodigo.Tag = Val(tCodigo.Text)
            If RsAux("VaCon") <> 0 Then
                MsgBox "El envío está en un VA CON el nuevo no será incluido en dicho VA CON", vbInformation, "Atención"
            End If
            lbDireccion.Caption = objGral.ArmoDireccionEnTexto(cBase, RsAux("EnvDireccion"))
            vsArticulos.Tag = RsAux("EnvFModificacion")
        End If
        RsAux.Close
    End If
        
    Cons = "Select Sum(REvAEntregar) as QArt, ArtID, ArtCodigo, rTrim(ArtNombre) as ArtNombre " & _
            " From RenglonEnvio, Articulo " & _
            " Where REvEnvio = " & Val(tCodigo.Tag) & _
            " And RevArticulo = ArtID And RevAEntregar > 0" & _
            " Group by ArtID, ArtCodigo, ArtNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    'Cargo la lista por si selecciona la opción EntregaParcial.
    Do While Not RsAux.EOF
        With vsArticulos
            .AddItem RsAux!QArt
            .Cell(flexcpText, .Rows - 1, 2) = "0"
            .Cell(flexcpText, .Rows - 1, 1) = "(" & Format(RsAux!ArtCodigo, "000,000") & ") " & Trim(RsAux!ArtNombre)
            .Cell(flexcpBackColor, .Rows - 1, 0, , 1) = vbWindowBackground
            lAux = RsAux!ArtID
            .Cell(flexcpData, .Rows - 1, 0) = lAux
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    Toolbar1.Buttons("save").Enabled = (vsArticulos.Rows > 1)
    On Error Resume Next
    If vsArticulos.Rows > 1 Then vsArticulos.SetFocus
    Screen.MousePointer = 0
    Exit Sub
errFE:
    Screen.MousePointer = 0
    vsArticulos.Rows = 1
    tCodigo.Tag = ""
    objGral.OcurrioError "Error al buscar el envío.", Err.Description
End Sub

Private Sub actSave()
On Error GoTo errSave
Dim iQ As Integer
Dim bQuedan As Boolean, bHay As Boolean
Dim rsEnv As rdoResultset, rsNew As rdoResultset
    
    lbMsg.Caption = "Almacenando"
    
    With vsArticulos
        For iQ = 1 To .Rows - 1
            If Val(.Cell(flexcpText, iQ, 0)) <> Val(.Cell(flexcpText, iQ, 2)) Then
                bQuedan = True
            End If
            If Val(.Cell(flexcpText, iQ, 2)) > 0 Then bHay = True
            If bHay And bQuedan Then Exit For
        Next
    End With
    If Not bQuedan Then
        MsgBox "Debe dejar artículos en el envío, no está dividiendo el envío.", vbExclamation, "Atención"
        Exit Sub
    End If
    If Not bHay Then
        MsgBox "No hay artículos seleccionados para el nuevo envío.", vbExclamation, "Atención"
        Exit Sub
    End If
    
    
    If MsgBox("¿Confirma dividir el envío?" & vbCrLf & vbCrLf & "El nuevo envío tendrá el mismo código de impresión debe darle un estado.", vbQuestion + vbYesNo, "Dividir el envío") = vbNo Then Exit Sub
    
    'Empiezo a copiar
    Screen.MousePointer = 11
    FechaDelServidor
    
    On Error GoTo errBT
    cBase.CommitTrans
    On Error GoTo ErrTransaccion
    
    Cons = "Select * From Envio where EnvCodigo = " & Val(tCodigo.Tag)
    Set rsEnv = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If rsEnv("EnvFModificacion") <> CDate(vsArticulos.Tag) Then
        rsEnv.Close
        cBase.RollbackTrans
        MsgBox "El envío fue modificado por otra terminal, cargue nuevamente la información.", vbExclamation, "Atención"
        Exit Sub
    End If
    
    Dim lNewDir As Long
    If Not IsNull(rsEnv("EnvDireccion")) Then
        lNewDir = CopioDireccion(rsEnv("EnvDireccion"))
    End If
    
    Cons = "Select * From Envio Where EnvCodigo = 0"
    Set rsNew = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    rsNew.AddNew
    rsNew("EnvTipo") = rsEnv("EnvTipo")
    rsNew("EnvDocumento") = rsEnv("EnvDocumento")
    If Not IsNull(rsEnv("EnvFechaPrometida")) Then rsNew("EnvFechaPrometida") = rsEnv("EnvFechaPrometida")
    If Not IsNull(rsEnv("EnvRangoHora")) Then rsNew("EnvRangoHora") = rsEnv("EnvRangoHora")
    If Not IsNull(rsEnv("EnvTipoFlete")) Then rsNew("EnvTipoFlete") = rsEnv("EnvTipoFlete")
    If Not IsNull(rsEnv("EnvCamion")) Then rsNew("EnvCamion") = rsEnv("EnvCamion")
    If Not IsNull(rsEnv("EnvAgencia")) Then rsNew("EnvAgencia") = rsEnv("EnvAgencia")
    If Not IsNull(rsEnv("EnvZona")) Then rsNew("EnvZona") = rsEnv("EnvZona")
    If Not IsNull(rsEnv("EnvFechaEntregado")) Then rsNew("EnvFechaEntregado") = rsEnv("EnvFechaEntregado")
    If Not IsNull(rsEnv("EnvCliente")) Then rsNew("EnvCliente") = rsEnv("EnvCliente")
    If Not IsNull(rsEnv("EnvComentario")) Then rsNew("EnvComentario") = rsEnv("EnvComentario")
'    If Not IsNull(rsEnv("EnvReclamoCobro")) Then rsNew("EnvReclamoCobro") = rsEnv("EnvReclamoCobro")
    
    If Not IsNull(rsEnv("EnvTelefono")) Then rsNew("EnvTelefono") = rsEnv("EnvTelefono")
           
    
    
    If Not IsNull(rsEnv("EnvMoneda")) Then rsNew("EnvMoneda") = rsEnv("EnvMoneda")
    If Not IsNull(rsEnv("EnvEstado")) Then rsNew("EnvEstado") = rsEnv("EnvEstado")
'    If Not IsNull(rsEnv("EnvLiquidar")) Then rsNew("EnvLiquidar") = rsEnv("EnvLiquidar")
'    If Not IsNull(rsEnv("EnvVolumenTotal")) Then rsNew("EnvVolumenTotal") = rsEnv("EnvVolumenTotal")
    
    If Not IsNull(rsEnv("EnvLiquidacion")) Then rsNew("EnvLiquidacion") = rsEnv("EnvLiquidacion")
 '   If Not IsNull(rsEnv("EnvBulto")) Then rsNew("EnvBulto") = rsEnv("EnvBulto")
'    If Not IsNull(rsEnv("EnvTamañoMayor")) Then rsNew("EnvTamañoMayor") = rsEnv("EnvTamañoMayor")
    If Not IsNull(rsEnv("EnvTipoHorario")) Then rsNew("EnvTipoHorario") = rsEnv("EnvTipoHorario")
    If Not IsNull(rsEnv("EnvCodImpresion")) Then rsNew("EnvCodImpresion") = rsEnv("EnvCodImpresion")
    
    'el documento que pago el envío lo dejo sólo en este.
    'rsNew("EnvDocumentoFactura") = null
    rsNew("EnvUsuario") = paCodigoDeUsuario
    rsNew("EnvFModificacion") = Format(Now, "yyyy/mm/dd hh:nn:ss")
    rsNew("EnvFormaPago") = 3       'le pongo factura camión
    If lNewDir > 0 Then rsNew("EnvDireccion") = lNewDir
    
    
    'If Not IsNull(rsEnv("EnvValorFlete")) Then rsNew("EnvValorFlete") = rsEnv("EnvValorFlete")
    'If Not IsNull(rsEnv("EnvIvaFlete")) Then rsNew("EnvIvaFlete") = rsEnv("EnvIvaFlete")

    'If Not IsNull(rsEnv("EnvValorPiso")) Then rsNew("EnvValorPiso") = rsEnv("EnvValorPiso")
    'If Not IsNull(rsEnv("EnvIvaPiso")) Then rsNew("EnvIvaPiso") = rsEnv("EnvIvaPiso")
        
    rsNew.Update
    rsNew.Close
    
    
    Cons = "Select Max(EnvCodigo) From Envio Where EnvTipo = " & rsEnv("EnvTipo") & " And EnvDocumento = " & rsEnv("EnvDocumento")
    Set rsNew = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    lNewDir = rsNew(0)
    rsNew.Close
    
'cambio la fecha de edición del viejo
    rsEnv.Edit
    rsEnv("EnvUsuario") = paCodigoDeUsuario
    rsEnv("EnvFModificacion") = Format(Now, "yyyy/mm/dd hh:nn:ss")
    rsEnv.Close
        
    With vsArticulos
        'Inserto los nuevos artículos
        For iQ = 1 To .Rows - 1
            If Val(.Cell(flexcpText, iQ, 2)) > 0 Then
                'Dos movimientos uno al viejo para restar y otro al nuevo
                If Val(.Cell(flexcpText, iQ, 0)) - Val(.Cell(flexcpText, iQ, 2)) > 0 Then
                    'Update
                    Cons = "Update RenglonEnvio Set REvCantidad = REvCantidad - " & Val(.Cell(flexcpText, iQ, 2)) & _
                        ", REvAEntregar = REvAEntregar - " & Val(.Cell(flexcpText, iQ, 2)) & _
                        " Where REvEnvio = " & Val(tCodigo.Tag) & " And REvArticulo = " & .Cell(flexcpData, iQ, 0)
                    cBase.Execute Cons
                    
                    Cons = "Insert INTO RenglonEnvio (REvEnvio, REvArticulo, REvCantidad, REvAEntregar) Values (" & _
                            lNewDir & ", " & .Cell(flexcpData, iQ, 0) & ", " & Val(.Cell(flexcpText, iQ, 2)) & ", " & Val(.Cell(flexcpText, iQ, 2)) & ")"
                    cBase.Execute Cons
                Else
                    'Al renglon del envío viejo le pongo el nuevo.
                    Cons = "Update RenglonEnvio Set REvEnvio = " & lNewDir & " Where REvEnvio = " & Val(tCodigo.Tag) & " And REvArticulo = " & .Cell(flexcpData, iQ, 0)
                    cBase.Execute Cons
                End If
                
            End If
        Next
    End With
    cBase.CommitTrans
    Screen.MousePointer = 0
    MsgBox "El código del nuevo envío es " & lNewDir & vbCrLf & vbCrLf & "Al mismo no se le asigno valor de flete, si es necesario hacerlo una vez que le eliminé el estado impreso podrá editarlo por el formulario de envíos.", vbInformation, "Atención"
    tCodigo.Text = ""
    vsArticulos.Rows = 1
    Exit Sub
        
errBT:
    Screen.MousePointer = vbDefault
    objGral.OcurrioError "Error al intentar iniciar la transacción para dividir el envío.", Err.Description, "Dividir envíos"
    Exit Sub

errorET:
    Resume ErrTransaccion
    Exit Sub
    
ErrTransaccion:
    Screen.MousePointer = vbDefault
    cBase.RollbackTrans
    objGral.OcurrioError "Error al grabar cuando se dividía el envío.", Err.Description, "Dividir envíos"
    Exit Sub
    
errSave:
    objGral.OcurrioError "Error al intentar al dividir el envío.", Err.Description, "Dividir envíos"
End Sub

Private Sub Form_Load()
    With vsArticulos
        .Rows = 1
        .FixedRows = 1
        .FormatString = "Q en envío|Artículo|Q para nuevo"
        .RowHeight(0) = 315
        .ColWidth(1) = 3500
        .BackColorSel = vbInfoBackground
        .ForeColorSel = vbWindowText
    End With
    Toolbar1.Buttons("save").Enabled = False
    If prmEnvio > 0 Then
        tCodigo.Text = prmEnvio
        loc_FindEnvio
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
    vsArticulos.Left = 0
    vsArticulos.Width = ScaleWidth
End Sub

Private Sub Label1_Click()
    With tCodigo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tCodigo_Change()
    If Val(tCodigo.Tag) > 0 Then vsArticulos.Rows = 1: Toolbar1.Buttons("save").Enabled = False: vsArticulos.Tag = ""
End Sub

Private Sub tCodigo_GotFocus()
    With tCodigo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    lbMsg.Caption = "Ingrese el código del envío que quiere dividir y presione enter para buscarlo."
End Sub

Private Sub tCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Not IsNumeric(tCodigo.Text) Then
            MsgBox "No es un código válido.", vbExclamation, "Atención"
        Else
            loc_FindEnvio
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "save": actSave
        Case "exit": Unload Me
    End Select
End Sub

Private Sub vsArticulos_GotFocus()
    lbMsg.Caption = "Seleccione el artículo que desea agregar al nuevo envío y presione + para quitarle al actual y darselo al nuevo."
End Sub

Private Sub vsArticulos_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errkd
    If Shift <> 0 Then Exit Sub
    With vsArticulos
        Select Case KeyCode
            Case vbKeyAdd
                If Val(.Cell(flexcpText, .Row, 2)) < Val(.Cell(flexcpText, .Row, 0)) Then .Cell(flexcpText, .Row, 2) = Val(.Cell(flexcpText, .Row, 2)) + 1
            Case vbKeySubtract
                If Val(.Cell(flexcpText, .Row, 2)) > 0 Then .Cell(flexcpText, .Row, 2) = Val(.Cell(flexcpText, .Row, 2)) - 1
        End Select
    End With
errkd:
End Sub

Private Sub vsArticulos_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn And Toolbar1.Buttons("save").Enabled Then actSave
End Sub
