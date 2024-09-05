VERSION 5.00
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form frmEstadoServicio 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6780
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picEstado 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   6750
      TabIndex        =   15
      Top             =   0
      Width           =   6780
      Begin VB.CheckBox chEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Con acces. y manual"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado del producto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   60
         Width           =   1815
      End
   End
   Begin VB.PictureBox picServicio 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   0
      ScaleHeight     =   1770
      ScaleWidth      =   6750
      TabIndex        =   13
      Top             =   1095
      Width           =   6780
      Begin VB.ComboBox cboQVias 
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtAclaracion 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   480
         Width           =   3975
      End
      Begin VB.TextBox txtMotivo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   840
         Width           =   2535
      End
      Begin AACombo99.AACombo cboRepararEn 
         Height          =   315
         Left            =   4200
         TabIndex        =   2
         Top             =   120
         Width           =   2535
         _ExtentX        =   4471
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
      Begin VSFlex6DAOCtl.vsFlexGrid lstMotivos 
         Height          =   870
         Left            =   3720
         TabIndex        =   17
         Top             =   840
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   1535
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
         BackColorSel    =   13686989
         ForeColorSel    =   0
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
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
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
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Cant. &Vías:"
         Height          =   255
         Left            =   5160
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Reparar en:"
         Height          =   255
         Left            =   3240
         TabIndex        =   1
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Aclaración:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Motivo:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Ingreso a servicio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.PictureBox picbotones 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   6750
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2895
      Width           =   6780
      Begin VB.CommandButton butAcciones 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5640
         TabIndex        =   10
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton butAcciones 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   4560
         TabIndex        =   9
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblAyuda 
         BackStyle       =   0  'Transparent
         Caption         =   "Devolución de mercaderí"
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
         TabIndex        =   12
         Top             =   60
         Width           =   4260
      End
   End
End
Attribute VB_Name = "frmEstadoServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oEstadoIngMerc As clsEstadosIngMerc
Public TipoDeEntrada As TipoAccionEntrada

Public IDArticulo As Long
Public Servicio As clsServicio
'Public Estado As Long
Public EstadosSeleccionados As String

Private Sub MostrarServicio(ByVal habilitados As Boolean)
    
    Dim lcolor As Long

    lcolor = IIf(habilitados, vbWindowBackground, vbButtonFace)
    
    With cboRepararEn
        .Enabled = habilitados
        .BackColor = lcolor
    End With
    
    With cboQVias
        .Enabled = habilitados
        .BackColor = lcolor
    End With
    
    With txtAclaracion
        .Enabled = habilitados
        .BackColor = lcolor
    End With
    
    With txtMotivo
        .Enabled = habilitados
        .BackColor = lcolor
    End With
    
    With lstMotivos
        .Enabled = habilitados
        .BackColor = lcolor
    End With

End Sub

Private Function ObtenerIDsSeleccionados() As String
Dim iQ As Byte
Dim ids As String
    For iQ = chEstado.LBound To chEstado.UBound
        If chEstado(iQ).value Then
            ids = ids & IIf(ids = "", "", ",") & chEstado(iQ).Tag
        End If
    Next
    ObtenerIDsSeleccionados = ids
End Function

Private Sub EsParaServicio()
Dim ids As String
     ids = ObtenerIDsSeleccionados
     MostrarServicio oEstadoIngMerc.EstadoEsARecuperar(ids)
End Sub

Private Sub CargoEstadosEnContenedor()
    
    If oEstadoIngMerc Is Nothing Then Exit Sub
    
    If oEstadoIngMerc.Estados.Count = 0 Then
        picEstado.Height = 0
    Else
        Dim oEst As clsEstadoMercaderia
        For Each oEst In oEstadoIngMerc.Estados
            If chEstado(0).Tag <> "" Then
                Load chEstado(chEstado.UBound + 1)
            End If
            With chEstado(chEstado.UBound)
            'Posiciono top y left
                Select Case (chEstado.UBound Mod 3)
                    Case 0
                        .Left = 120
                    Case 1
                        .Left = 2280
                    Case 2
                        .Left = 4560
                End Select
                .Top = 360 + ((chEstado.UBound \ 3) * 240)
                .Tag = oEst.ID
                .TabIndex = chEstado.UBound
                .Caption = oEst.Nombre
                .Visible = True
                picEstado.Height = .Top + 240
            End With
        Next
        picEstado.Height = picEstado.Height + 60
    End If
    Me.Height = picbotones.Height + picbotones.Top
    
End Sub

Private Sub butAcciones_Click(Index As Integer)
    
    Select Case Index
        Case 0
            Dim ids As String
            ids = ObtenerIDsSeleccionados()
            If ids = "" Then
                MsgBox "Debe seleccionar al menos uno de los estados.", vbExclamation, "Validación"
                chEstado(0).SetFocus
                Exit Sub
            End If
            
            If cboRepararEn.Enabled Then
                If cboRepararEn.ListIndex = -1 Then
                    MsgBox "Indique si se repara el artículo.", vbExclamation, "Validación"
                    cboRepararEn.SetFocus
                    Exit Sub
                End If
                
                If cboRepararEn.ItemData(cboRepararEn.ListIndex) = 0 Then
                    If cboQVias.ListIndex = -1 Then
                        MsgBox "Seleccione la cantidad de vías que desea imprimir.", vbExclamation, "Validación"
                        cboQVias.SetFocus
                        Exit Sub
                    End If
                    
                Else
                    
                    If cboQVias.ListIndex = -1 Then
                        MsgBox "Seleccione la cantidad de vías que desea imprimir.", vbExclamation, "Validación"
                        cboQVias.SetFocus
                        Exit Sub
                    End If
                
                End If
                
                If cboQVias.Text <> "0" Then
                    If Trim(txtAclaracion.Text) = "" Then
                        MsgBox "Ingrese un comentario para la ficha.", vbExclamation, "Validación"
                        txtAclaracion.SetFocus
                        Exit Sub
                    End If
                Else
                    If cboRepararEn.ItemData(cboRepararEn.ListIndex) > 0 Then
                        Dim result As VbMsgBoxResult
                        Do
                            result = MsgBox("¿Confirma no imprimir las fichas de servicio?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Impresión de fichas")
                        Loop Until result <> vbCancel
                        If result = vbNo Then
                            cboQVias.SetFocus
                            Exit Sub
                        End If
                    End If
                End If
            End If
            
            'Obtengo el valor de los items seleccionados.
            'Estado = oEstadoIngMerc.ObtenerValorEstadoSeleccionado( ids)
            EstadosSeleccionados = ObtenerIDsSeleccionados
            
            'Si tiene servicio entonces cargo la clase
            Set Servicio = Nothing
            Set Servicio = New clsServicio
            
            If cboRepararEn.Enabled Then

                With Servicio

                    .LocalRepara.ID = cboRepararEn.ItemData(cboRepararEn.ListIndex)
                    .LocalRepara.Nombre = cboRepararEn.Text
                    .Aclaracion = txtAclaracion.Text
                    .Vias = Val(cboQVias.Text)
                    
                    If lstMotivos.Rows > 0 Then
                        Dim iQ As Integer
                        Dim oMotivo As clsCodigoTexto
                        For iQ = 0 To lstMotivos.Rows - 1
                            Set oMotivo = New clsCodigoTexto
                            oMotivo.ID = lstMotivos.Cell(flexcpData, iQ, 0)
                            oMotivo.Nombre = lstMotivos.Cell(flexcpText, iQ, 0)
                            Servicio.Motivos.Add oMotivo
                        Next
                    End If

                End With
            
            End If
            Unload Me
            
        Case 1
            EstadosSeleccionados = ""
            Unload Me
            
    End Select
    
End Sub


Private Sub cboQVias_GotFocus()
    
    lblAyuda.Caption = "Seleccione la cantidad de fichas que desea imprimir."
    If cboQVias.ListIndex = -1 Then
        If cboRepararEn.ListIndex < 1 Then
            cboQVias.Text = "0"
        Else
            cboQVias.Text = "1"
        End If
    End If
    
End Sub

Private Sub cboQVias_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtMotivo.SetFocus
End Sub

Private Sub cboQVias_LostFocus()
    lblAyuda.Caption = ""
End Sub

Private Sub cboRepararEn_GotFocus()

    With cboRepararEn
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    lblAyuda.Caption = "Si el artículo requiere ingreso a taller seleccione en que local se reparará."

End Sub

Private Sub cboRepararEn_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cboRepararEn.ListIndex > -1 Then txtAclaracion.SetFocus
End Sub

Private Sub cboRepararEn_LostFocus()
    lblAyuda.Caption = ""
End Sub

Private Sub chEstado_Click(Index As Integer)
    EsParaServicio
End Sub

Private Sub chEstado_GotFocus(Index As Integer)
    'frmEntMercaderia.MostrarAyuda
    lblAyuda.Caption = "Indique los estados que considere aplica el estado en que recibe el artículo."
End Sub

Public Sub LimpiarControles()
    
    Dim iQ As Integer
    For iQ = chEstado.LBound To chEstado.UBound
        chEstado(iQ).value = 0
    Next
    
    cboRepararEn.Text = ""
    cboQVias.Text = "0"
    txtAclaracion.Text = ""
    txtMotivo.Text = ""
    lstMotivos.Rows = 0
    
End Sub

Private Sub chEstado_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If cboRepararEn.Enabled And cboRepararEn.Visible Then
            cboRepararEn.SetFocus
        Else
            butAcciones(0).SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        EstadosSeleccionados = ""
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    
    lstMotivos.Rows = 0
    
    cboQVias.Clear
    cboQVias.AddItem "0"
    cboQVias.AddItem "1"
    cboQVias.AddItem "2"
    
    
    If chEstado(0).Tag = "" Then
        
        CargoEstadosEnContenedor
    
        'Cargo Sucursales---------------------------------------------------------------------------
        cboRepararEn.Clear
        Dim oLocal As clsCodigoTexto
        For Each oLocal In colLocalesRepara
            With cboRepararEn
                .AddItem oLocal.Nombre
                .ItemData(.NewIndex) = oLocal.ID
            End With
        Next
        '-----------------------------------------------------------------------------------------------
    
    End If
    
    Dim iQ As Integer, iQ1 As Byte
    If EstadosSeleccionados <> "" Then
    
        Dim vIDs() As String
        vIDs = Split(oEstadoIngMerc.ObtenerCadenaEstadosSeleccionados(EstadosSeleccionados), ",")
        If UBound(vIDs) >= 0 Then
            For iQ = 0 To UBound(vIDs)
                If Trim(vIDs(iQ)) <> "" Then
                    For iQ1 = chEstado.LBound To chEstado.UBound
                        If Trim(chEstado(iQ1).Caption) = Trim(vIDs(iQ)) Then
                            chEstado(iQ1).value = 1
                            Exit For
                        End If
                    Next
                End If
            Next
        End If
        
    End If
    
    If Not Servicio Is Nothing And cboRepararEn.Enabled Then
        With Servicio
            txtAclaracion.Text = .Aclaracion
            
            For iQ = 0 To cboRepararEn.ListCount
                If cboRepararEn.ItemData(iQ) = .LocalRepara.ID Then
                    cboRepararEn.ListIndex = iQ
                    Exit For
                End If
            Next
            
            cboQVias.ListIndex = .Vias
            
            If .Motivos.Count > 0 Then
                Dim oMotivo As clsCodigoTexto
                For iQ = 1 To .Motivos.Count
                    Set oMotivo = .Motivos(iQ)
                    lstMotivos.AddItem oMotivo.Nombre
                    lstMotivos.Cell(flexcpData, lstMotivos.Rows - 1, 0) = oMotivo.ID
                Next
            End If
            
        End With
    End If
    
    
End Sub

Private Sub lstMotivos_GotFocus()
    lblAyuda.Caption = "Motivos del servicio. (Supr borra el renglón)"
End Sub

Private Sub lstMotivos_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If lstMotivos.Rows = 0 Then Exit Sub
    If KeyCode = vbKeyDelete Then
        lstMotivos.RemoveItem lstMotivos.Row
    End If
End Sub

Private Sub lstMotivos_LostFocus()
    lblAyuda.Caption = ""
End Sub

Private Sub txtAclaracion_GotFocus()
    
    lblAyuda.Caption = "Ingrese un comentario para la ficha de servicio."
    With txtAclaracion
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub txtAclaracion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(txtAclaracion.Text) <> "" Then cboQVias.SetFocus
End Sub

Private Sub txtAclaracion_LostFocus()
    lblAyuda.Caption = ""
End Sub


Private Sub txtMotivo_GotFocus()
    With txtMotivo
        
        If .Text = "" And cboRepararEn.ListIndex >= 1 Then .Text = "%"
        If .Text = "%" Then .SelStart = Len(.Text): Exit Sub
        
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    lblAyuda.Caption = "Ingrese parte o el nombre del motivo y de enter para buscar."
End Sub

Private Sub txtMotivo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
                
        If Trim(txtMotivo.Text) = "" Then
            
            On Error Resume Next
            butAcciones(0).SetFocus
        
        Else
            On Error GoTo ErrBM
            Screen.MousePointer = 11
                
            Cons = "Select MSeID, Nombre = MSeNombre From MotivoServicio " _
                & " Where MSeTipo = (Select ArtTipo From Articulo Where ArtID = " & IDArticulo & ")" _
                & " And MSeNombre Like '" & Replace(txtMotivo.Text, " ", "%") & "%'"
            
            Dim objHelp As New clsListadeAyuda
            If objHelp.ActivarAyuda(cBase, Cons, 3000, 1, "Lista de motivos") > 0 Then
                'recorro la lista para ver si ya lo ingreso.
                Dim iQ As Byte
                If lstMotivos.Rows > 0 Then
                    For iQ = 0 To lstMotivos.Rows - 1
                        If Val(lstMotivos.Cell(flexcpData, iQ, 0)) = objHelp.RetornoDatoSeleccionado(0) Then
                            MsgBox "El motivo seleccionado ya está ingresado.", vbExclamation, "Duplicación"
                            Screen.MousePointer = 0
                            Exit Sub
                        End If
                    Next
                End If
                lstMotivos.AddItem objHelp.RetornoDatoSeleccionado(1)
                lstMotivos.Cell(flexcpData, lstMotivos.Rows - 1, 0) = objHelp.RetornoDatoSeleccionado(0)
            End If
            Set objHelp = Nothing
            Screen.MousePointer = 0
            
            txtMotivo.Text = "": txtMotivo.Tag = ""
            
        End If
    End If
    Exit Sub
ErrBM:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error al buscar los motivos.", Trim(Err.Description)

End Sub

Private Sub txtMotivo_LostFocus()
    lblAyuda.Caption = ""
End Sub
