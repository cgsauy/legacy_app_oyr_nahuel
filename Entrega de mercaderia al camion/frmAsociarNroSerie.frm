VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form frmAsociarNroSerie 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asociar número de serie"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtArticulo 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   1560
      Width           =   6255
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsArticulos 
      Height          =   2415
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4260
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
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
   Begin VB.CommandButton butCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6480
      TabIndex        =   9
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton butAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox txtNroSerie 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2040
      Width           =   6255
   End
   Begin VB.TextBox txtRemito 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1080
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   8085
      TabIndex        =   0
      Top             =   0
      Width           =   8115
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Height          =   510
         Left            =   120
         Picture         =   "frmAsociarNroSerie.frx":0000
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   1
         Top             =   120
         Width           =   510
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Asociar número de serie a remito"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Artículo:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblArticulo 
      Caption         =   "Nro. de serie:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblRemito 
      Caption         =   "Remito"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Remito"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "frmAsociarNroSerie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oRenglon As Collection

Sub GrabarDatos()

    If vsArticulos.Rows = 1 Then
        MsgBox "No hay artículos para grabar.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    If Val(txtRemito.Tag) = 0 Then
        MsgBox "No hay un remito seleccionado.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    If MsgBox("¿Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
        GraboProductosVendidos
    End If
    
End Sub

Private Sub GraboProductosVendidos()
    On Error GoTo errBT
    Screen.MousePointer = 11
    cBase.BeginTrans
    
    Cons = "DELETE ProductosVendidos WHERE PVeDocumento = " & Val(txtRemito.Tag)
    cBase.Execute Cons
    
    Dim nroSerie As Variant
    
    Dim oArt As clsRenglon
    For Each oArt In oRenglon
        If oArt.NumerosDeSerie.Count > 0 Then
            
            For Each nroSerie In oArt.NumerosDeSerie
                
                Cons = "Select * FROM ProductosVendidos " & _
                    " Where PVeArticulo = " & oArt.Articulo & " AND PVeDocumento = " & Val(txtRemito.Tag)
                
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                RsAux.AddNew
                RsAux("PVeDocumento") = Val(txtRemito.Tag)
                RsAux("PVeArticulo") = Val(oArt.Articulo)
                RsAux("PVeNSerie") = nroSerie
                RsAux("PVeVarGarantia") = 1
                RsAux.Update
                RsAux.Close
                
            Next
        End If
    Next
    
    cBase.CommitTrans
    On Error GoTo errYaGrabe
    
    txtRemito.Text = ""
    Screen.MousePointer = 0
    Exit Sub

errBT:
    objGral.OcurrioError "Error inesperado al inicializar la transacción.", Err.Description, "Grabar"
    Screen.MousePointer = 0
    Exit Sub
    
    
errYaGrabe:
    objGral.OcurrioError "Error inesperado al finalizar el evento grabar.", Err.Description, "Restauración de formulario"
    Screen.MousePointer = 0
    Exit Sub
    
ErrResumo:
    Resume ErrRelajo
    
ErrRelajo:
    Screen.MousePointer = vbDefault
    cBase.RollbackTrans
    objGral.OcurrioError "Ocurrió un error al emitir los remitos de cambio.", Err.Description, "Grabar"
    Exit Sub

End Sub

Private Sub butAceptar_Click()
    GrabarDatos
End Sub

Private Sub butCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    With vsArticulos
        .Rows = 1
        .Cols = 1
        .FormatString = "Artículo|Serie"
        .ColWidth(0) = 3500
        .ExtendLastCol = True
    End With
    OcultoEntrada
End Sub

Sub OcultoEntrada()
    lblRemito.Caption = ""
    txtRemito.Text = ""
    txtArticulo.Text = ""
    txtNroSerie.Text = ""
    txtArticulo.Enabled = False
    txtNroSerie.Enabled = False
    txtArticulo.BackColor = vbApplicationWorkspace
    txtNroSerie.BackColor = vbApplicationWorkspace
    vsArticulos.Rows = 1
    vsArticulos.Enabled = False
    vsArticulos.Tag = 0
End Sub

Private Sub txtArticulo_Change()
    If Val(txtArticulo.Tag) > 0 Then txtArticulo.Tag = "": txtNroSerie.Text = ""
End Sub

Private Sub txtArticulo_GotFocus()
    With txtArticulo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(txtArticulo.Tag) > 0 Then
            txtNroSerie.SetFocus
        ElseIf txtArticulo.Text = "" Then
            butAceptar.SetFocus
        Else
            BuscoArticuloEscaneado
        End If
    End If
End Sub

Private Sub txtNroSerie_GotFocus()
    With txtNroSerie
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtNroSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtNroSerie.Text <> "" Then
            If (Val(txtArticulo.Tag) > 0) Then
                InsertarSerieArticulo
            Else
                txtNroSerie.Text = ""
                txtArticulo.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtRemito_Change()
    If Val(txtRemito.Tag) > 0 Then
        lblRemito.Caption = ""
        txtRemito.Tag = ""
        Set oRenglon = Nothing
        txtArticulo.Text = ""
        txtNroSerie.Text = ""
        txtArticulo.Enabled = False
        txtNroSerie.Enabled = False
        txtArticulo.BackColor = vbApplicationWorkspace
        txtNroSerie.BackColor = vbApplicationWorkspace
        vsArticulos.Rows = 1
        vsArticulos.Enabled = False
        vsArticulos.Tag = 0
    End If
End Sub

Private Sub txtRemito_GotFocus()
    With txtRemito
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtRemito_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(txtRemito.Tag) > 0 Then
            txtArticulo.SetFocus
        Else
            BuscoRemito
        End If
    End If
End Sub

Private Sub BuscoArticuloEscaneado()
On Error GoTo errBAE
Dim sQy As String, sCodBar As String
Dim RsAux As rdoResultset
Dim iRetQ As Integer
Dim esCombo As Boolean

    iRetQ = 1
    txtNroSerie.Tag = ""
    sCodBar = Replace(txtArticulo.Text, "'", "''")
    sQy = "EXEC prg_BuscarArticuloEscaneado '" + sCodBar + "'"
    Set RsAux = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        
        If Not IsNull(RsAux("ArtID")) Then
        
            txtArticulo.Text = RsAux("ArtNombre")
            txtArticulo.Tag = RsAux("ArtID")
            
            If Trim(sCodBar) <> Trim(RsAux("ArtCodigo")) And RsAux("ACBLargo") > 0 Then
                txtNroSerie.Text = sCodBar
                'INSERTO EL Artículo en la grilla.
                InsertarSerieArticulo
            Else
                Dim bExiste As Boolean
                Dim oArt As clsRenglon
                For Each oArt In oRenglon
                    If oArt.Articulo = Val(txtArticulo.Tag) Then
                        bExiste = True
                        Exit For
                    End If
                Next
                If Not bExiste Then
                    RsAux.Close
                    MsgBox "El artículo escaneado no pertenece al remito.", vbExclamation, "ATENCIÓN"
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                If RsAux("PedirNSerie") Then txtNroSerie.Tag = "s"
            End If
        End If
    End If
    RsAux.Close
    
    If Val(txtArticulo.Tag) > 0 Then
        txtNroSerie.SetFocus
    Else
        txtArticulo.SetFocus
    End If
    Exit Sub
    
errBAE:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al buscar el artículo.", Err.Description, "Buscar remito"
End Sub

Private Function InsertarSerieArticulo() As Boolean
    Dim oArt As clsRenglon
    For Each oArt In oRenglon
        If oArt.Articulo = Val(txtArticulo.Tag) Then
            If oArt.SerieDuplicada(txtNroSerie.Text) Then
                MsgBox "ATENCIÓN, serie duplicada.", vbExclamation, "POSIBLE ERROR"
                Exit Function
            Else
                If oArt.NumerosDeSerie.Count < oArt.ARetirar Then
                    InsertarFila txtArticulo.Text, oArt.Articulo, txtNroSerie.Text
                    oArt.NumerosDeSerie.Add txtNroSerie.Text
                    InsertarSerieArticulo = True
                    If oArt.NumerosDeSerie.Count = oArt.ARetirar Then
                        txtArticulo.Text = ""
                        txtArticulo.SetFocus
                    End If
                    Exit Function
                Else
                    MsgBox "La cantidad de artículos escaneados supera los del remito.", vbExclamation, "ATENCIÓN"
                    txtArticulo.Text = ""
                    Exit Function
                End If
            End If
        End If
    Next
    MsgBox "El artículo escaneado no pertenece al remito.", vbExclamation, "ATENCIÓN"
End Function

Private Sub InsertarFila(ByVal nomArticulo As String, ByVal idArticulo As Long, ByVal serie As String)
    With vsArticulos
        .AddItem nomArticulo
        .Cell(flexcpData, .Rows - 1, 0) = idArticulo
        .Cell(flexcpText, .Rows - 1, 1) = serie
    End With
End Sub

Private Sub BuscoRemito()
On Error GoTo errBR
    Dim sQy As String
    Dim rsD As rdoResultset
    Dim texto As String
    
    Set oRenglon = New Collection
    texto = Trim(txtRemito.Text)
    vsArticulos.Tag = 0
    Dim Codigo As Long
    Dim TipoDoc As TipoDocumento
    Codigo = FormatoBarras(txtRemito.Text, TipoDoc)
    If Codigo = 0 Then
        texto = Replace(Replace(texto, "-", ""), " ", "")
        If Len(texto) < 2 Then
            MsgBox "Formato incorrecto.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
        'buscoRemito por serie y nro.
        sQy = "WHERE DocTipo = " & TipoDocumento.RemitoEntrega & " AND DocSerie = '" & Mid(texto, 1, 1) & "' AND DocNumero = " & Mid(texto, 2)
    Else
        If TipoDoc = RemitoEntrega Then
            sQy = "WHERE DocCodigo = " & Codigo
        Else
            MsgBox "Sólo se admiten remitos de entrega de mercadería.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
    
    End If
    Screen.MousePointer = 11
    
    sQy = "SELECT DocCodigo, DocSerie, DocNumero, SucNombre, RenArticulo, RenCantidad, RenCantidad FROM Documento INNER JOIN Sucursal ON DocSucursal = SucCodigo " & _
            "INNER JOIN Renglon ON RenDocumento = DocCodigo AND RenCantidad > 0 " & _
            "INNER JOIN Envio ON EnvDocumento = DocCodigo AND EnvEstado <= 3 " & _
            sQy

    Set rsD = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    If rsD.EOF Then
    
        MsgBox "No hay un remito con los datos ingresados o el mismo no tiene artículos pendientes a entregar.", vbExclamation, "Buscar remito"
        OcultoEntrada
        
    Else
        
        txtRemito.Text = ""
        txtRemito.Text = Trim(rsD("DocSerie")) & "-" & rsD("DocNumero")
        
        lblRemito.Caption = Trim(rsD("SucNombre")) & " " & Trim(rsD("DocSerie")) & "-" & rsD("DocNumero")
        txtRemito.Tag = rsD("DocCodigo")
        
        txtArticulo.BackColor = vbWindowBackground
        txtNroSerie.BackColor = vbWindowBackground
        txtArticulo.Enabled = True
        txtNroSerie.Enabled = True
        vsArticulos.Enabled = True
        
        Dim rsP As rdoResultset
        
        Dim oFila As clsRenglon
        Do While Not rsD.EOF
            Set oFila = New clsRenglon
            
            oRenglon.Add oFila
            With oFila
                .ARetirar = rsD("RenCantidad")
                .Articulo = rsD("RenArticulo")
                Set .NumerosDeSerie = New Collection
            End With
            
            Cons = "SELECT ArtNombre, PVeArticulo, PVeNSerie FROM ProductosVendidos INNER JOIN Articulo ON ArtID = PVeArticulo WHERE PVeDocumento = " & Val(txtRemito.Tag) & _
                " AND PVeArticulo = " & rsD("RenArticulo")
            Set rsP = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
            Do While Not rsP.EOF
                vsArticulos.Tag = 1
                oFila.NumerosDeSerie.Add Trim(CStr(rsP("PVeNSerie")))
                InsertarFila Trim(rsP("ArtNombre")), rsP("PVeArticulo"), rsP("PVeNSerie")
                rsP.MoveNext
                
            Loop
            rsP.Close
            
            rsD.MoveNext
        Loop
        
    End If
    rsD.Close
    
    Screen.MousePointer = 0
    Exit Sub
    
errBR:
Screen.MousePointer = 0
    objGral.OcurrioError "Error al buscar el remito.", Err.Description, "Buscar remito"
End Sub

Private Function FormatoBarras(ByVal texto As String, ByRef TipoDoc As TipoDocumento) As Long
Dim iDBarCode As String
    On Error GoTo errInt
    texto = UCase(texto)
    FormatoBarras = 0
    If InStr(1, texto, "D", vbTextCompare) > 1 Then
        If Not IsNumeric(Mid(texto, 1, InStr(texto, "D") - 1)) Then Exit Function
        If IsNumeric(Mid(texto, 1, InStr(texto, "D") - 1)) And IsNumeric(Trim(Mid(texto, InStr(texto, "D") + 1, Len(texto)))) Then
            TipoDoc = CLng(Mid(texto, 1, InStr(texto, "D") - 1))
            FormatoBarras = CLng(Trim(Mid(texto, InStr(texto, "D") + 1, Len(texto))))
        End If
    End If
    Exit Function
errInt:
    Screen.MousePointer = 0
    objGral.OcurrioError "Ocurrió un error en el formato de barras.", Err.Description
End Function

Private Sub vsArticulos_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error Resume Next
    If KeyCode = vbKeyDelete And vsArticulos.Rows > 1 Then
        'Lo busco en la collección y los saco a los 2.
        Dim nroSerie As Variant
        Dim oArt As clsRenglon
        For Each oArt In oRenglon
            If oArt.Articulo = Val(vsArticulos.Cell(flexcpData, vsArticulos.RowSel, 0)) Then
                Dim iR As Integer
                For iR = 1 To oArt.NumerosDeSerie.Count
                'For Each nroSerie In oArt.NumerosDeSerie
                    If oArt.NumerosDeSerie.Item(iR) = vsArticulos.Cell(flexcpText, vsArticulos.RowSel, 1) Then
                        oArt.NumerosDeSerie.Remove iR
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next
        vsArticulos.RemoveItem vsArticulos.RowSel
    End If
    
End Sub

