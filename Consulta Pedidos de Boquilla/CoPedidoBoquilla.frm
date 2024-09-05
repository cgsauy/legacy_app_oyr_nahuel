VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.0#0"; "AACOMBO.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form CoPedidoBoquilla 
   Caption         =   "Consulta de Pedidos de Boquilla Pendientes"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   9180
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CoPedidoBoquilla.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   5115
      TabIndex        =   19
      Top             =   4080
      Width           =   5175
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "CoPedidoBoquilla.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Picture         =   "CoPedidoBoquilla.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "CoPedidoBoquilla.frx":06D6
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "CoPedidoBoquilla.frx":0820
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   4320
         Picture         =   "CoPedidoBoquilla.frx":096A
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   3600
         Picture         =   "CoPedidoBoquilla.frx":0A6C
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3240
         Picture         =   "CoPedidoBoquilla.frx":0E32
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   2295
      Left            =   4080
      TabIndex        =   21
      Top             =   2040
      Width           =   4215
      _ExtentX        =   7435
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   12
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
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   5565
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   7990
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fFiltros 
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1455
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   9015
      Begin VB.TextBox tArticulo 
         Height          =   315
         Left            =   5400
         TabIndex        =   7
         Top             =   600
         Width           =   3375
      End
      Begin AACombo99.AACombo cGrupo 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
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
         Text            =   ""
      End
      Begin AACombo99.AACombo cTipo 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
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
         Text            =   ""
      End
      Begin AACombo99.AACombo cMarca 
         Height          =   315
         Left            =   5400
         TabIndex        =   3
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
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
         Text            =   ""
      End
      Begin AACombo99.AACombo cProveedor 
         Height          =   315
         Left            =   1080
         TabIndex        =   9
         Top             =   960
         Width           =   3135
         _ExtentX        =   5530
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
         Text            =   ""
      End
      Begin VB.Label Label5 
         Caption         =   "&Proveedor:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "&Marca:"
         Height          =   255
         Left            =   4680
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "&Tipo:"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "&Grupo:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   4680
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
   End
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   3495
      Left            =   240
      TabIndex        =   20
      Top             =   1920
      Visible         =   0   'False
      Width           =   8895
      _Version        =   196608
      _ExtentX        =   15690
      _ExtentY        =   6165
      _StockProps     =   229
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
      PreviewMode     =   1
      Zoom            =   70
      EmptyColor      =   0
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin VB.Menu MnuAccesos 
      Caption         =   "Accesos"
      Visible         =   0   'False
      Begin VB.Menu MnuPedidoBoquilla 
         Caption         =   "Pedidos de boquilla"
      End
      Begin VB.Menu MnuEliminarPedido 
         Caption         =   "Eliminar Pedido"
      End
   End
End
Attribute VB_Name = "CoPedidoBoquilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const PorPantalla = 200
Private PosLista As Integer
Private aSeleccionado As Long

Private RsConsulta As rdoResultset

Private aFormato As String, aTituloTabla As String, aComentario As String

Private aTexto As String

Private Sub AccionSiguiente()

    On Error GoTo ErrAA
    If Not bSiguiente.Enabled Then Exit Sub
    
    If Not RsConsulta.EOF Then
        bAnterior.Enabled = True: bPrimero.Enabled = True
        
        PosLista = PosLista + (vsConsulta.Rows - 1)
        CargoLista
    Else
        MsgBox "Se ha llegado al final de la consulta, no hay más datos a desplegar.", vbInformation, "ATENCIÓN"
    End If
    
    vsConsulta.SetFocus
    Exit Sub
    
ErrAA:
    msgError.MuestroError "Ocurrió un error inesperado. " & Err.Description
End Sub

Private Sub AccionAnterior()

Dim UltimaPosicion As Long

    On Error GoTo ErrAA
    If Not bAnterior.Enabled Then Exit Sub
    
    If RsConsulta.EOF And vsConsulta.Rows - 1 > 0 And RsConsulta.AbsolutePosition = -1 Then
        Screen.MousePointer = 11
        RsConsulta.MoveLast
        Screen.MousePointer = 0
        UltimaPosicion = PosLista + vsConsulta.Rows
        RsConsulta.MoveNext
        Screen.MousePointer = 0
    Else
        UltimaPosicion = PosLista + vsConsulta.Rows
    End If
    
    If UltimaPosicion - (vsConsulta.Rows - 1) - PorPantalla >= 1 Then
        
        If UltimaPosicion - (vsConsulta.Rows - 1) - PorPantalla = 1 Then bAnterior.Enabled = False: bPrimero.Enabled = False
        bSiguiente.Enabled = True
        
        RsConsulta.Move UltimaPosicion - (vsConsulta.Rows - 1) - PorPantalla, 1
        CargoLista
        PosLista = PosLista - (vsConsulta.Rows - 1)
        Screen.MousePointer = 0
    Else
        MsgBox "Se ha llegado al principio de la consulta.", vbInformation, "ATENCIÓN"
    End If
    vsConsulta.SetFocus
    Exit Sub
    
ErrAA:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error inesperado. " & Err.Description
End Sub

Private Sub AccionPrimero()
    
    If Not bPrimero.Enabled Then Exit Sub
    
    PosLista = 0
    Screen.MousePointer = 11
    On Error Resume Next
    RsConsulta.MoveFirst
    On Error GoTo ErrAA
    CargoLista
    vsConsulta.SetFocus
    bSiguiente.Enabled = True
    bPrimero.Enabled = False: bAnterior.Enabled = False
    Screen.MousePointer = 0
    Exit Sub
ErrAA:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error inesperado. " & Err.Description
End Sub
Private Sub AccionLimpiar()
    cGrupo.Text = ""
    cTipo.Text = ""
    cMarca.Text = ""
    cProveedor.Text = ""
    tArticulo.Text = ""
End Sub
Private Sub bAnterior_Click()
    AccionAnterior
End Sub
Private Sub bCancelar_Click()
    Unload Me
End Sub
Private Sub bConsultar_Click()
    AccionConsultar
End Sub
Private Sub bImprimir_Click()
    AccionImprimir
End Sub
Private Sub bNoFiltros_Click()
    AccionLimpiar
End Sub
Private Sub bPrimero_Click()
    AccionPrimero
End Sub
Private Sub bSiguiente_Click()
    AccionSiguiente
End Sub
Private Sub cGrupo_GotFocus()
    cGrupo.SelStart = 0
    cGrupo.SelLength = Len(cGrupo.Text)
End Sub
Private Sub cGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco tArticulo
End Sub

Private Sub cMarca_GotFocus()
    With cMarca
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub cMarca_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cGrupo.SetFocus
End Sub

Private Sub cMarca_LostFocus()
    cGrupo.SelStart = 0
End Sub

Private Sub cProveedor_GotFocus()
    With cProveedor
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub cProveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub cProveedor_LostFocus()
    cProveedor.SelStart = 0
End Sub

Private Sub cTipo_GotFocus()
    With cTipo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub cTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco cMarca
End Sub

Private Sub cTipo_LostFocus()
    cTipo.SelStart = 0
End Sub

Private Sub Label2_Click()
    Foco cTipo
End Sub

Private Sub Label4_Click()
    Foco cMarca
End Sub

Private Sub Label5_Click()
    Foco cProveedor
End Sub

Private Sub MnuEliminarPedido_Click()
On Error GoTo ErrEP
    
    If MsgBox("¿Confirma eliminar el pedido nro. " & CLng(vsConsulta.Cell(flexcpText, vsConsulta.Row, 0)) & "?", vbQuestion + vbYesNo, "ELIMINAR PEDIDO") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    Cons = "Select * From Pedido Where PedCodigo = " & CLng(vsConsulta.Cell(flexcpText, vsConsulta.Row, 0))
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No se encontró el pedido, otra terminal pudo eliminarla o modificarla, verifique.", vbExclamation, "ATENCIÓN"
    Else
        RsAux.Delete
        RsAux.Close
    End If
    vsConsulta.RemoveItem vsConsulta.Row
    vsConsulta.Refresh
    Screen.MousePointer = 0
    Exit Sub
ErrEP:
    msgError.MuestroError "Ocurrio un error al eliminar el pedido."
    Screen.MousePointer = 0
End Sub

Private Sub MnuPedidoBoquilla_Click()
On Error GoTo errApp
    Dim RetVal
    Dim Parametro As String
    Screen.MousePointer = 11
    Parametro = vsConsulta.Cell(flexcpText, vsConsulta.Row, 0)
    RetVal = Shell(App.Path & "\Pedido de Boquilla " & Parametro, 1)
    Screen.MousePointer = 0
    Exit Sub
errApp:
    msgError.MuestroError "Ocurrió un error al ejecutar la aplicación. " & Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = 0
End Sub
Private Sub tArticulo_GotFocus()
    tArticulo.SelStart = 0
    tArticulo.SelLength = Len(tArticulo.Text)
End Sub
Private Sub tArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrAK
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyF1 Then
    
        If Val(tArticulo.Tag) = 0 And Trim(tArticulo.Text) <> "" Then
        
            If Not IsNumeric(tArticulo.Text) Then   'Busqueda de articulos por lista de ayuda-------------------
                Screen.MousePointer = 11
                Dim aLista As New clsListadeAyuda
                Cons = "Select ArtID, 'Descripción' = ArtNombre, 'Código' = ArtCodigo From Articulo " _
                        & " Where ArtNombre Like '" & Trim(tArticulo.Text) & "%'" _
                        & " Order by ArtNombre"
                aLista.ActivoListaAyuda Cons, False, cBase.Connect
                
                aSeleccionado = aLista.ValorSeleccionado
                If aSeleccionado <> 0 Then
                    tArticulo.Text = aLista.ItemSeleccionado
                    tArticulo.Tag = aSeleccionado
                    bConsultar.SetFocus
                Else
                    tArticulo.Text = ""
                End If
                
                Set aLista = Nothing
                Screen.MousePointer = 0
            
            Else    'Busqueda de Articulos por codigo--------------
                Screen.MousePointer = 11
                Cons = "Select ArtID, ArtNombre from Articulo where ArtCodigo = " & Trim(tArticulo.Text)
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                If Not RsAux.EOF Then
                    tArticulo.Text = Trim(RsAux!ArtNombre)
                    tArticulo.Tag = RsAux!ArtID
                    bConsultar.SetFocus
                Else
                    MsgBox "No existe un artículo para el código ingresado.", vbExclamation, "ATENCIÓN"
                End If
                RsAux.Close
                Screen.MousePointer = 0
            End If
            
        Else
             Foco cProveedor
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrAK:
    msgError.MuestroError "Ocurrio un error inesperado.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
Private Sub Form_Activate()
    Screen.MousePointer = 0: Me.Refresh
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad
    
    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    
    bPrimero.Enabled = False: bSiguiente.Enabled = False: bAnterior.Enabled = False
    LimpioGrilla
    
    Cons = "Select * from Articulo Where ArtID = 0"
    Set RsConsulta = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    'Cargo Combos.---------------------------------
    'Tipo.
    Cons = "Select TipCodigo, TipNombre From Tipo Order by TipNombre"
    CargoCombo Cons, cTipo
    'Marca
    Cons = "Select MarCodigo, MarNombre From Marca Order by MarNombre"
    CargoCombo Cons, cMarca
    'Grupos.
    Cons = "Select GruCodigo, GruNombre From Grupo Order by GruNombre"
    CargoCombo Cons, cGrupo
    'Proveedor
    Cons = "Select PExCodigo, PExNombre From ProveedorExterior Order by PExNombre"
    CargoCombo Cons, cProveedor
    '------------------------------------------------------
    FechaDelServidor
    gFechaServidor = Format(gFechaServidor, FormatoFP)
    AccionLimpiar
    Exit Sub
ErrLoad:
    msgError.MuestroError "Ocurrio un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyE: AccionConsultar
            
            Case vbKeyP: AccionPrimero
            Case vbKeyA: AccionAnterior
            Case vbKeyS: AccionSiguiente
            
            Case vbKeyQ: AccionLimpiar
            Case vbKeyI: AccionImprimir
            
            Case vbKeyX: Unload Me
        End Select
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    Screen.MousePointer = 11
    picBotones.BorderStyle = vbFlat
    picBotones.Top = Me.ScaleHeight - (picBotones.Height + Status.Height + 40)
    fFiltros.Width = Me.Width - (fFiltros.Left * 2.5)
    vsConsulta.Left = fFiltros.Left
    vsConsulta.Top = fFiltros.Top + fFiltros.Height + 50
    vsConsulta.Height = Me.ScaleHeight - (vsConsulta.Top + picBotones.Height + Status.Height + 90)
    vsConsulta.Width = fFiltros.Width
    Screen.MousePointer = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    GuardoSeteoForm Me
    RsConsulta.Close
    CierroConexion
    Set msgError = Nothing
    End
End Sub

Private Sub Label1_Click()
    Foco tArticulo
End Sub

Private Sub Label3_Click()
    Foco cGrupo
End Sub

Private Sub AccionConsultar()
    
    If Not VerificoFiltros Then Exit Sub
    
    LimpioGrilla
    
    'Cierro el cursor.---------------------------------
    On Error Resume Next: RsConsulta.Close
    
    On Error GoTo errConsultar
    
    Screen.MousePointer = 11
    PosLista = 0
    ArmoConsulta
    Set RsConsulta = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsConsulta.EOF Then
        bPrimero.Enabled = False: bAnterior.Enabled = False: bSiguiente.Enabled = False
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        Screen.MousePointer = 0
        Exit Sub
    End If

    CargoLista
    
    If RsConsulta.EOF Then bPrimero.Enabled = False: bAnterior.Enabled = False: bSiguiente.Enabled = False
    
    vsConsulta.ColSel = 0
    vsConsulta.ColSort(0) = flexSortGenericAscending
    vsConsulta.Sort = flexSortUseColSort
    
    Screen.MousePointer = 0
    Exit Sub

errConsultar:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al realizar la consulta de datos.", Err.Description
End Sub

Private Sub ArmoConsulta()

    Cons = "Select PedCodigo, PedFPedido, PedFEmbarque, PExNombre, AFoPUnitario, AFoCantidad, ArtNombre " _
        & " From Pedido, ArticuloFolder, Articulo, ProveedorExterior" _
        & " Where PedCarpeta Is Null" _
        & " And ArtHabilitado = 'S' And ArtSeImporta = 1 "
        
    If Val(tArticulo.Tag) > 0 Then Cons = Cons & " And ArtID = " & tArticulo.Tag
    If cTipo.ListIndex > -1 Then Cons = Cons & " And ArtTipo = " & cTipo.ItemData(cTipo.ListIndex)
    If cMarca.ListIndex > -1 Then Cons = Cons & " And ArtMarca = " & cMarca.ItemData(cMarca.ListIndex)
    If cProveedor.ListIndex > -1 Then Cons = Cons & " And ArtProveedor = " & cProveedor.ItemData(cProveedor.ListIndex)
    
    If cGrupo.ListIndex > -1 Then
        Cons = Cons & " And ArtID IN (" _
                & " Select AGrArticulo from ArticuloGrupo" _
                & " Where AGrGrupo = " & cGrupo.ItemData(cGrupo.ListIndex) & ")"
    End If
    
    Cons = Cons & " And AFoTipo = " & Folder.cFPedido _
        & " And PedCodigo = AFoCodigo  And AFoArticulo = ArtID" _
        & " And PedProveedor = PExCodigo "
    
End Sub

Private Sub CargoLista()

    On Error GoTo ErrInesperado
    
    vsConsulta.Redraw = False
    vsConsulta.Rows = 1
    
    Screen.MousePointer = 11
    Do While Not RsConsulta.EOF And vsConsulta.Rows - 1 < PorPantalla
         
        'Inserto en la grilla.------------------------------------------------
        vsConsulta.AddItem "", vsConsulta.Rows
        With vsConsulta
            .Cell(flexcpText, vsConsulta.Rows - 1, 0) = Trim(RsConsulta!PedCodigo)
            .Cell(flexcpText, vsConsulta.Rows - 1, 1) = Format(RsConsulta!PedFpedido, "yyyy-mm-dd")
            .Cell(flexcpText, vsConsulta.Rows - 1, 2) = Trim(RsConsulta!PExNombre)
            .Cell(flexcpText, vsConsulta.Rows - 1, 3) = Format(RsConsulta!PedFEmbarque, "yyyy-mm-dd")
            .Cell(flexcpText, vsConsulta.Rows - 1, 4) = Trim(RsConsulta!ArtNombre)
            .Cell(flexcpText, vsConsulta.Rows - 1, 5) = RsConsulta!AFoCantidad
            .Cell(flexcpText, vsConsulta.Rows - 1, 6) = RsConsulta!AFoPUnitario
        End With
        RsConsulta.MoveNext
    Loop
    
    If RsConsulta.EOF Then bSiguiente.Enabled = False Else bSiguiente.Enabled = True
    vsConsulta.Redraw = True
    Screen.MousePointer = 0
    Exit Sub
    
ErrInesperado:
    vsConsulta.Redraw = True
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al realizar la consulta de stock.", Err.Description
End Sub

Private Sub AccionImprimir()
Dim J As Integer

    If vsConsulta.Rows = 1 Then
        MsgBox "No hay datos en la lista para realizar la impresión.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    On Error GoTo errPrint
    Screen.MousePointer = 11
    aTituloTabla = "": aComentario = ""
    
    With vsListado
        If Not .PrintDialog(pdPrinterSetup) Then Screen.MousePointer = 0: Exit Sub
        
        .Preview = True
        .StartDoc
                
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión.", vbInformation, "ATENCIÓN"
            Screen.MousePointer = vbDefault: Exit Sub
        End If
    
        EncabezadoListado vsListado, "Importaciones - Consulta de Pedidos de Boquilla.", False
        
        .filename = "Consulta de Pedidos de Boquilla"
        .FontSize = 8: .FontBold = False
                
        For I = 1 To vsConsulta.Rows - 1
            aTexto = Trim(vsConsulta.TextMatrix(I, 0))
            For J = 1 To vsConsulta.Cols - 1
                aTexto = aTexto & "|" & Trim(vsConsulta.TextMatrix(I, J))
            Next J
            .AddTable aFormato, "", aTexto, Colores.inactivo, , True
        Next I
        
        .EndDoc
        .PrintDoc
    End With
    
    Screen.MousePointer = 0
    Exit Sub

errPrint:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al realizar la impresión. ", Err.Description
End Sub

Private Function ArmoFormulaFiltros() As String

Dim aRetorno As String

    On Error Resume Next
    aRetorno = ""
    
    If cTipo.ListIndex <> -1 Then aRetorno = aRetorno & " Tipo: " & cTipo.Text & ", "
    If cMarca.ListIndex <> -1 Then aRetorno = aRetorno & " Marca: " & cMarca.Text & ", "
    If cGrupo.ListIndex <> -1 Then aRetorno = aRetorno & " Grupo: " & cGrupo.Text & ", "
    If Val(tArticulo.Tag) <> 0 Then aRetorno = aRetorno & "Art.: " & Trim(tArticulo.Text) & ", "
    If cProveedor.ListIndex <> -1 Then aRetorno = aRetorno & " Prov.: " & cProveedor.Text & ", "
    
    aRetorno = Mid(aRetorno, 1, Len(aRetorno) - 2)
    ArmoFormulaFiltros = aRetorno

End Function

Private Function VerificoFiltros() As Boolean

    VerificoFiltros = False
    
    If tArticulo.Text <> "" And Val(tArticulo.Tag) = 0 Then
        MsgBox "El artículo seleccionado no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tArticulo: Exit Function
    End If
    
    If cGrupo.Text <> "" And cGrupo.ListIndex = -1 Then
        MsgBox "El grupo de artículos ingresado no es correcto.", vbExclamation, "ATENCIÓN"
        Foco cGrupo: Exit Function
    End If
    
    If cTipo.Text <> "" And cTipo.ListIndex = -1 Then
        MsgBox "El tipo de artículos no es correcto.", vbExclamation, "ATENCIÓN"
        Foco cTipo: Exit Function
    End If
    
    If cMarca.Text <> "" And cMarca.ListIndex = -1 Then
        MsgBox "La marca de artículos no es correcta.", vbExclamation, "ATENCIÓN"
        Foco cMarca: Exit Function
    End If
    
    If cProveedor.Text <> "" And cProveedor.ListIndex = -1 Then
        MsgBox "El proveedor de artículos no es correcto.", vbExclamation, "ATENCIÓN"
        Foco cProveedor: Exit Function
    End If
    
    VerificoFiltros = True

End Function

Private Sub ImpresionEncabezadoTabla()

    aTituloTabla = "": aFormato = ""
    
    For I = 0 To vsConsulta.Cols - 1
        Select Case vsConsulta.ColAlignment(I)
            Case lvwColumnCenter: aFormato = aFormato & "+^~"
            Case lvwColumnLeft: aFormato = aFormato & "+<~"
            Case lvwColumnRight: aFormato = aFormato & "+>~"
        End Select
        aFormato = aFormato & CInt(vsConsulta.ColWidth(I) * 1.5) & "|"
        aTituloTabla = aTituloTabla & vsConsulta.TextMatrix(0, I) & "|"
    Next
    
    aFormato = Mid(aFormato, 1, Len(aFormato) - 1)
    aTituloTabla = Mid(aTituloTabla, 1, Len(aTituloTabla) - 1)
        
End Sub
Private Sub vsConsulta_Click()
    If vsConsulta.MouseRow = 0 Then
        vsConsulta.ColSel = vsConsulta.MouseCol
        If vsConsulta.ColSort(vsConsulta.MouseCol) = flexSortGenericAscending Then
            vsConsulta.ColSort(vsConsulta.MouseCol) = flexSortGenericDescending
        Else
            vsConsulta.ColSort(vsConsulta.MouseCol) = flexSortGenericAscending
        End If
        vsConsulta.Sort = flexSortUseColSort
    End If
End Sub

Private Sub vsConsulta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And vsConsulta.Rows > 1 Then PopupMenu MnuAccesos, X:=vsConsulta.Left + X, Y:=Y + vsConsulta.Top
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub vsListado_NewPage()

    If aTituloTabla = "" Then
        ImpresionEncabezadoTabla
        aComentario = "Filtros: " & ArmoFormulaFiltros
    End If
    
    With vsListado
        .Paragraph = aComentario
        .FontSize = 8: .FontBold = True
        .TableBorder = tbBoxRows
        .AddTable aFormato, aTituloTabla, "", , Colores.inactivo
        .FontBold = False
    End With
    
End Sub
Private Sub LimpioGrilla()

    With vsConsulta
        .Redraw = False
        .ExtendLastCol = True
        .Clear
        .Editable = True
        .Rows = 1
        .Cols = 7
        .FormatString = "Codigo|^Fecha|Proveedor|^Embarca|Articulo|Q|Costo"
        .ColWidth(0) = 0
        .ColWidth(1) = 1200
        .ColWidth(2) = 1500
        .ColWidth(3) = 1200
        .ColWidth(4) = 1800
        .ColWidth(5) = 600
        .ColWidth(6) = 950
        .ColHidden(0) = True
        .AllowUserResizing = flexResizeColumns
        .MergeCells = flexMergeRestrictAll
        .MergeCol(0) = True: .MergeCol(1) = True: .MergeCol(2) = True: .MergeCol(3) = True
        .Redraw = True
    End With

End Sub
