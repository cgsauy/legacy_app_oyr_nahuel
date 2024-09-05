VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Begin VB.Form frmListado 
   BackColor       =   &H00C0C000&
   Caption         =   "Ingreso de Existencia"
   ClientHeight    =   6840
   ClientLeft      =   3360
   ClientTop       =   2610
   ClientWidth     =   9870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   9870
   Begin VB.ComboBox cGrupo 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   5940
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox cQue 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.PictureBox picFin 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   9870
      TabIndex        =   30
      Top             =   5355
      Width           =   9870
      Begin VB.TextBox tUsuario 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   32
         Top             =   420
         Width           =   615
      End
      Begin VB.TextBox tComentario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   31
         Top             =   60
         Width           =   6315
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "&Usuario:"
         Height          =   255
         Left            =   180
         TabIndex        =   35
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "C&omentario:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   60
         Width           =   975
      End
      Begin VB.Label labUsuario 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1740
         TabIndex        =   33
         Top             =   420
         Width           =   975
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   3075
      Left            =   300
      TabIndex        =   12
      Top             =   2220
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5424
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
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
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
      OutlineBar      =   1
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
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   3135
      Left            =   60
      TabIndex        =   27
      Top             =   1560
      Width           =   6075
      _Version        =   196608
      _ExtentX        =   10716
      _ExtentY        =   5530
      _StockProps     =   229
      BackColor       =   -2147483633
      BorderStyle     =   1
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
      Zoom            =   70
      EmptyColor      =   0
      AbortWindowPos  =   0
      AbortWindowPos  =   0
      BackColor       =   -2147483633
   End
   Begin VB.PictureBox picBotones 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C000&
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9810
      TabIndex        =   28
      Top             =   6150
      Width           =   9870
      Begin VB.CommandButton bConexion 
         Caption         =   "CB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   5400
         Picture         =   "frmListado.frx":08CA
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Cambiar Base de Datos."
         Top             =   60
         Width           =   375
      End
      Begin VB.CommandButton bModificar 
         Height          =   310
         Left            =   480
         Picture         =   "frmListado.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Nuevo. [Ctrl+N]"
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   1260
         Picture         =   "frmListado.frx":0B16
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Cancelar. [Ctrl+C]"
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bGrabar 
         Height          =   310
         Left            =   900
         Picture         =   "frmListado.frx":0C18
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Grabar. [Ctrl+G]"
         Top             =   60
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0D1A
         Height          =   310
         Left            =   4380
         Picture         =   "frmListado.frx":0E1C
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   3540
         Picture         =   "frmListado.frx":134E
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   3180
         Picture         =   "frmListado.frx":1438
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   2760
         Picture         =   "frmListado.frx":1522
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   4020
         Picture         =   "frmListado.frx":175C
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bSalir 
         Height          =   310
         Left            =   4860
         Picture         =   "frmListado.frx":185E
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bNuevo 
         Height          =   310
         Left            =   120
         Picture         =   "frmListado.frx":1960
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Nuevo. [Ctrl+N]"
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   2040
         Picture         =   "frmListado.frx":1A62
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   2400
         Picture         =   "frmListado.frx":1DA4
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   1680
         Picture         =   "frmListado.frx":20A6
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   60
         Width           =   310
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   6585
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   14340
            TextSave        =   ""
            Key             =   "msg"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fFiltros 
      BackColor       =   &H00C0C000&
      Caption         =   "Ingreso de datos."
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
      Left            =   60
      TabIndex        =   25
      Top             =   0
      Width           =   8715
      Begin VB.CommandButton bComArticulo 
         Caption         =   "Co&mentario ..."
         Height          =   315
         Left            =   5340
         TabIndex        =   36
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox tCantidad 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   900
         TabIndex        =   11
         Top             =   960
         Width           =   4275
      End
      Begin VB.TextBox tArticulo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   900
         TabIndex        =   7
         Top             =   600
         Width           =   4275
      End
      Begin AACombo99.AACombo cEstado 
         Height          =   315
         Left            =   6000
         TabIndex        =   9
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
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
      Begin AACombo99.AACombo cLocal 
         Height          =   315
         Left            =   900
         TabIndex        =   1
         Top             =   240
         Width           =   2115
         _ExtentX        =   3731
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
         BackStyle       =   0  'Transparent
         Caption         =   "&Cantidad:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "&Grupo:"
         Height          =   255
         Left            =   5340
         TabIndex        =   4
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Qué:"
         Height          =   255
         Left            =   3360
         TabIndex        =   2
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   660
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Estado:"
         Height          =   255
         Left            =   5340
         TabIndex        =   8
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Local:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   495
      End
   End
   Begin ComctlLib.ImageList ImgList 
      Left            =   5820
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmListado.frx":22E0
            Key             =   "Alta"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmListado.frx":25FA
            Key             =   "Baja"
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuPUComArt 
      Caption         =   "ComentarioArticulo"
      Visible         =   0   'False
      Begin VB.Menu MnuPUCAInComentario 
         Caption         =   "Ingreso de Comentario"
      End
      Begin VB.Menu MnuPUACViewComentario 
         Caption         =   "Ver Comentarios"
      End
      Begin VB.Menu MnuL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEstado 
         Caption         =   "MnuEstado"
         Index           =   0
      End
      Begin VB.Menu MnuL2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "Cancelar"
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private aTexto As String
Private sNuevo As Boolean

Private Sub bCancelar_Click()
    AccionCancelar
End Sub

Private Sub bComArticulo_Click()
    frmComentario.prmArticulo = Val(tArticulo.Tag)
    frmComentario.Show vbModal, Me
End Sub

Private Sub bConexion_Click()
Dim newB As String
    
    On Error GoTo errCh
    
    If Not miConexion.AccesoAlMenu("Cambiar_Conexion") Then Exit Sub
    
    newB = InputBox("Ingrese el texto del login para la nueva conexión" & vbCrLf & _
                "Id de aplicación en archivo de conexiones.", "Cambio de Base de Datos")
    
    If Trim(newB) = "" Then Exit Sub
    If MsgBox("Está seguro de cambiar la base de datos al login " & newB, vbQuestion + vbYesNo + vbDefaultButton2, "Cambiar Base") = vbNo Then Exit Sub
    
    'Limpio la ficha
    If bCancelar.Enabled Then AccionCancelar
    cLocal.Text = "": tComentario.Text = vbNullString: tComentario.Text = "": labUsuario.Caption = ""
    cQue.ListIndex = -1: cGrupo.ListIndex = -1
    
    newB = miConexion.TextoConexion(newB)
    With frmListado
        If Me.Tag = miConexion.RetornoPropiedad(bdb:=True) Then
            .BackColor = &HC0C000
        Else
            .BackColor = &HFFFFC0
        End If
        picBotones.BackColor = .BackColor
        fFiltros.BackColor = .BackColor
        picFin.BackColor = .BackColor
    End With
    If Trim(newB) = "" Then Exit Sub
    
    Screen.MousePointer = 11
    On Error Resume Next
    cBase.Close
    
    On Error GoTo errCh
    Set cBase = Nothing
    Set cBase = eBase.OpenConnection("", rdDriverNoPrompt, , newB)
    If InicioConexionBD(newB) Then
        Status.Panels("bd").Text = "BD: " & miConexion.RetornoPropiedad(bdb:=True) & "  "
    Else
        Status.Panels("bd").Text = "BD: EN ERROR  "
    End If
    Screen.MousePointer = 0
    MsgBox "Ahora está trabajanbo en la nueva base de datos.", vbExclamation, "Base Cambiada OK"
    Exit Sub
    
errCh:
    clsGeneral.OcurrioError "Error de Conexión." & vbCrLf & " La conexión está en estado de error, conectese a una base de datos.", Err.Description
    Status.Panels("bd").Text = "BD: EN ERROR  "
    Screen.MousePointer = 0
End Sub

Private Sub bGrabar_Click()
    AccionGrabar
End Sub

Private Sub bImprimir_Click()
    AccionImprimir True
End Sub

Private Sub bModificar_Click()
    AccionModificar
End Sub

Private Sub bNuevo_Click()
    AccionNuevo
End Sub

Private Sub bPrimero_Click()
    IrAPagina vsListado, 1
End Sub

Private Sub bSalir_Click()
    Unload Me
End Sub
Private Sub bSiguiente_Click()
    IrAPagina vsListado, vsListado.PreviewPage + 1
End Sub
Private Sub bUltima_Click()
    IrAPagina vsListado, vsListado.PageCount
End Sub

Private Sub bZMas_Click()
    Zoom vsListado, vsListado.Zoom + 5
End Sub
Private Sub bZMenos_Click()
    Zoom vsListado, vsListado.Zoom - 5
End Sub

Private Sub bAnterior_Click()
    IrAPagina vsListado, vsListado.PreviewPage - 1
End Sub

Private Sub cEstado_GotFocus()
    If cEstado.Text = "" Then BuscoCodigoEnCombo cEstado, CLng(paEstadoArticuloEntrega)
    With cEstado
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Seleccione el estado de la mercadería."
End Sub

Private Sub cEstado_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If cEstado.ListIndex <> -1 Then InsertoRenglon
    End If
    
End Sub

Private Sub cEstado_LostFocus()
    cEstado.SelStart = 0: Ayuda ""
End Sub

Private Sub cLocal_Click()
    vsConsulta.Rows = 1
End Sub

Private Sub cLocal_Change()
    vsConsulta.Rows = 1
End Sub
Private Sub cLocal_GotFocus()
    With cLocal
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Seleccione el local donde ingresará la mercadería."
End Sub

Private Sub cLocal_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        fnc_CargoCombosConteo
        Foco cQue
    End If
    
End Sub

Private Sub cLocal_LostFocus()
    cLocal.SelStart = 0
    Ayuda ""
End Sub

Private Sub chVista_Click()

    If chVista.Value = 0 Then
        vsConsulta.ZOrder 0
        Me.Refresh
    Else
        AccionImprimir False
        vsListado.ZOrder 0
        Me.Refresh
    End If
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0: Me.Refresh
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    picBotones.BorderStyle = vbBSNone
    FechaDelServidor
    Status.Panels("bd").Text = "BD: " & miConexion.RetornoPropiedad(bdb:=True) & "  "
    Me.Tag = miConexion.RetornoPropiedad(bdb:=True)
    
    cons = "Select LocCodigo, LocNombre From Local Order by LocNombre"
    CargoCombo cons, cLocal
    
    cons = "Select EsMCodigo, EsMAbreviacion From EstadoMercaderia Order by EsMAbreviacion"
    CargoCombo cons, cEstado
    
    Dim idX As Integer: idX = 0
    cons = "Select EsMCodigo, EsMAbreviacion From EstadoMercaderia Order by EsMAbreviacion"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not rsAux.EOF
        If idX > 0 Then Load MnuEstado(idX)
        With MnuEstado(idX)
            .Caption = Trim(rsAux!EsMAbreviacion)
            .Tag = Trim(rsAux!EsMCodigo)
        End With
        idX = idX + 1
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    
    
    sNuevo = True
    AccionCancelar
    sNuevo = False
    
    'Hoja Carta
    vsListado.Orientation = orPortrait: vsListado.PaperSize = 1: vsListado.BackColor = Blanco
    InicializoGrillas
    cLocal.Enabled = True: cLocal.BackColor = Obligatorio
    labUsuario.Caption = ""
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub
Private Sub InicializoGrillas()
    On Error Resume Next
    With vsConsulta
'        .Editable = True
        .Redraw = False
        .WordWrap = False
        .Cols = 1: .Rows = 1
        .FormatString = "Tipo|>Cantidad|<Artículo|<Estado|"
        .ColWidth(0) = 800: .ColWidth(1) = 800: .ColWidth(2) = 3700: .ColWidth(3) = 800
        .ColHidden(0) = True
        .Redraw = True
        
        .Editable = True
        
    End With
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Shift = vbCtrlMask Then
        Select Case KeyCode
            
            Case vbKeyP: IrAPagina vsListado, 1
            Case vbKeyA: IrAPagina vsListado, vsListado.PreviewPage - 1
            Case vbKeyS: IrAPagina vsListado, vsListado.PreviewPage + 1
            Case vbKeyU: IrAPagina vsListado, vsListado.PageCount
            
            Case vbKeyAdd: Zoom vsListado, vsListado.Zoom + 5
            Case vbKeySubtract: Zoom vsListado, vsListado.Zoom - 5
            
            Case vbKeyI: AccionImprimir True
                
            Case vbKeyL: If chVista.Value = 0 Then chVista.Value = 1 Else chVista.Value = 0
            Case vbKeyN: AccionNuevo
            Case vbKeyM: AccionModificar
            Case vbKeyC: AccionCancelar
            Case vbKeyG: AccionGrabar
            Case vbKeyX: Unload Me
        End Select
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If tUsuario.Enabled Then
        If MsgBox("Usted no Grabó !!!" & vbCrLf & vbCrLf & "¿Está seguro que quiere salir sin grabar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Screen.MousePointer = 11

    vsListado.Height = Me.ScaleHeight - (vsListado.Top + Status.Height + 70 + picBotones.Height + picFin.Height)
'    picBotones.Top = Status.Top - picBotones.Height   'vsListado.Height + vsListado.Top + 70
    
    fFiltros.Width = Me.ScaleWidth - (vsListado.Left * 2)
    vsListado.Width = fFiltros.Width
    vsListado.Left = fFiltros.Left
    
    vsConsulta.Top = vsListado.Top
    vsConsulta.Width = vsListado.Width
    vsConsulta.Height = vsListado.Height
    vsConsulta.Left = vsListado.Left
    
    Screen.MousePointer = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    GuardoSeteoForm Me
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
    
End Sub

Private Sub Label1_Click()
    Foco cLocal
End Sub
Private Sub Label2_Click()
    Foco cEstado
End Sub
Private Sub Label5_Click()
    Foco tCantidad
End Sub
Private Sub Label6_Click()
    Foco cGrupo
End Sub
Private Sub Label7_Click()
    Foco tComentario
End Sub
Private Sub Label8_Click()
    Foco tUsuario
End Sub

Private Sub MnuEstado_Click(Index As Integer)
On Error GoTo errCambiar
    
    Dim xRow As Integer
    With vsConsulta
        xRow = .Row
    
        .Cell(flexcpData, xRow, 2) = Val(MnuEstado(Index).Tag)
        .Cell(flexcpText, xRow, 3) = MnuEstado(Index).Caption
    End With
    
errCambiar:
End Sub

Private Sub MnuPUACViewComentario_Click()
    EjecutarApp App.Path & "\comentario_inventario.exe", CStr(vsConsulta.Cell(flexcpData, vsConsulta.Row, 1))
End Sub

Private Sub MnuPUCAInComentario_Click()
    frmComentario.prmArticulo = Val(vsConsulta.Cell(flexcpData, vsConsulta.Row, 1))
    frmComentario.Show vbModal, Me
End Sub

Private Sub cQue_Change()
    If sNuevo Then vsConsulta.Rows = 1
End Sub

Private Sub cQue_GotFocus()
    Ayuda ""
End Sub

Private Sub cQue_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cGrupo
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = 0
End Sub

Private Sub tArticulo_GotFocus()
    With tArticulo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese el mes y año de liquidación a consultar."
End Sub
Private Sub tArticulo_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrAP
    If KeyAscii = vbKeyReturn Then
        If Trim(tArticulo.Text) <> "" Then
            If Val(tArticulo.Tag) <> 0 Then Foco tCantidad: Exit Sub
            Screen.MousePointer = 11
            
            If IsNumeric(tArticulo.Text) Then
                BuscoArticuloPorCodigo tArticulo.Text
                If Val(tArticulo.Tag) = 0 Then BuscoArticuloPorNombre tArticulo.Text
            Else
                BuscoArticuloPorNombre tArticulo.Text
            End If
            If Val(tArticulo.Tag) > 0 Then Foco tCantidad
        Else
            If vsConsulta.Rows > 1 Then Foco tComentario
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrAP:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub tArticulo_LostFocus()
    Ayuda ""
End Sub

Private Sub tCantidad_GotFocus()
    With tCantidad: .SelStart = 0: .SelLength = Len(.Text): End With
    Ayuda "Ingrese la cantidad de artículos."
End Sub

Private Sub tCantidad_KeyPress(KeyAscii As Integer)
On Error GoTo errCant

    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tCantidad.Text) Then
            If Val(tCantidad.Text) <> 0 Then Foco cEstado
        Else
            tCantidad.Text = clsGeneral.Eval(tCantidad.Text)
            tCantidad.SelStart = Len(tCantidad.Text)
        End If
    End If
    
    Exit Sub
errCant:
    MsgBox "Error: " & Err.Description, vbCritical, "ATENCIÓN"
End Sub

Private Sub cGrupo_GotFocus()
    Ayuda ""
End Sub

Private Sub cGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cLocal.ListIndex <> -1 And cQue.ListIndex <> -1 And cGrupo.ListIndex <> -1 Then
            CargoBalance
            Foco tArticulo
        End If
    End If
End Sub

Private Sub tComentario_GotFocus()
    tComentario.SelStart = 0
    tComentario.SelLength = Len(tComentario.Text)
    Ayuda " Ingrese un comentario."
End Sub
Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tUsuario.SetFocus
End Sub
Private Sub tComentario_LostFocus()
    Ayuda ""
End Sub
Private Sub tUsuario_GotFocus()
    tUsuario.SelStart = 0: tUsuario.SelLength = Len(tUsuario.Text)
    Ayuda " Ingrese su código de Usuario."
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)
On Error GoTo errTU
    
    If KeyAscii = vbKeyReturn Then
        tUsuario.Tag = 0
        If IsNumeric(tUsuario.Text) Then
            tUsuario.Tag = z_BuscoUsuarioDigito(Val(tUsuario.Text), True)
            If Val(tUsuario.Tag) > 0 Then AccionGrabar Else tUsuario.Tag = vbNullString
        Else
            AccionGrabar
        End If
    End If
    Exit Sub
errTU:
    clsGeneral.OcurrioError "Error en Usuario. (KeyPress)", Err.Description
End Sub

Private Sub vsConsulta_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If Col <> 1 Then Cancel = True: Exit Sub
    If Not tCantidad.Enabled Then Cancel = True: Exit Sub
       
    
End Sub

Private Sub vsConsulta_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errKD
    If Not bGrabar.Enabled Then Exit Sub
    
Dim xRow As Integer, mMSG As String

    If (vsConsulta.Rows <= vsConsulta.FixedRows) Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyDelete
            xRow = vsConsulta.Row
            
            mMSG = Replace("¿Confirma dejar la cantidad de «[p]» en cero?", "[p]", vsConsulta.Cell(flexcpText, xRow, 2))
            If MsgBox(mMSG, vbQuestion + vbYesNo, "Cantidad a Cero") = vbYes Then
                vsConsulta.Cell(flexcpText, xRow, 1) = 0
            End If
    End Select
    

errKD:
End Sub

Private Sub vsConsulta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If vsConsulta.Rows > 1 And Button = vbRightButton Then
        Dim idX As Integer, xRow As Integer
        
        xRow = vsConsulta.Row
        For idX = MnuEstado.LBound To MnuEstado.UBound
            MnuEstado(idX).Enabled = bGrabar.Enabled
            MnuEstado(idX).Checked = (Val(vsConsulta.Cell(flexcpData, xRow, 2)) = Val(MnuEstado(idX).Tag))
        Next
        
        PopupMenu MnuPUComArt
    End If
End Sub

Private Sub vsConsulta_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If Not IsNumeric(vsConsulta.EditText) Then Cancel = True
    
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionImprimir(Imprimir As Boolean)
Dim Consulta As Boolean
    On Error GoTo errImprimir
    
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    If vsConsulta.Rows > 1 Then

        With vsListado
            .StartDoc
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        
        vsListado.FileName = "Ingreso de Mercadería de Existencia"
        EncabezadoListado vsListado, "Ingreso de Mercadería de Existencia al " & Format(gFechaServidor, FormatoFP), False
                
        If cLocal.ListIndex > -1 Then
            aTexto = "LOCAL: |" & cLocal.Text
            aTexto = aTexto & "| Qué: |" & cQue.Text
            aTexto = aTexto & "| Grupo: |" & cGrupo.Text
            vsListado.TableBorder = tbBox
            vsListado.AddTable ">800|1500|>800|1200|>1000|1200", "", Trim(aTexto)
            vsListado.Paragraph = ""
        End If
        
        If Trim(tComentario.Text) <> "" Then vsListado.Paragraph = "Comentario: " & Trim(tComentario.Text)
        If Val(tUsuario.Tag) > 0 Then
            vsListado.Paragraph = "Usuario: " & z_BuscoUsuario(tUsuario.Tag, True)
        Else
            If Trim(labUsuario.Caption) <> "" Then vsListado.Paragraph = "Usuario: " & Trim(labUsuario.Caption)
        End If
        vsListado.Paragraph = ""
        vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
        vsListado.EndDoc
        
    End If
    
    If Imprimir Then
        frmSetup.pControl = vsListado
        frmSetup.Show vbModal, Me
        Me.Refresh
        If frmSetup.pOK Then vsListado.PrintDoc , frmSetup.pPaginaD, frmSetup.pPaginaH
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errImprimir:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
End Sub

Private Sub AccionConfigurar()
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
End Sub

Private Sub Ayuda(strTexto As String)
    Status.Panels("msg").Text = strTexto
End Sub

Private Sub AccionGrabar()
Dim IdBML As Long
Dim sErr As String
    On Error GoTo errPru
    
    sErr = "válido grilla."
    If vsConsulta.Rows = 1 Then MsgBox "No hay artículos ingresados.", vbExclamation, "ATENCIÓN": Exit Sub
    sErr = "válido usuario."
    If Val(tUsuario.Tag) = 0 Then MsgBox "Debe ingresar su código de usuario.", vbExclamation, "ATENCIÓN": Foco tUsuario: Exit Sub
    
    sErr = "válido comentario."
    If Not clsGeneral.TextoValido(tComentario.Text) Then MsgBox "Se ingreso un carácter no válido en el comentario.", vbExclamation, "ATENCIÓN": Exit Sub
    sErr = "válido local."
    If cLocal.ListIndex = -1 Then MsgBox "Debe seleccionar un local.", vbExclamation, "ATENCIÓN": Foco cLocal: Exit Sub
    sErr = "válido área 2."
    If cQue.ListIndex = -1 Then MsgBox "Falta el ingreso de algunos datos obligatorios", vbExclamation, "ATENCIÓN": Foco cQue: Exit Sub
    sErr = "válido nro. ingreso."
    If cGrupo.ListIndex = -1 Then MsgBox "Debe ingresar un nro. de ingreso.", vbExclamation, "ATENCIÓN": Foco cGrupo: Exit Sub
    
    Dim xIDQue As Long, xIDGrupo As Long
    
    xIDQue = cQue.ItemData(cQue.ListIndex)
    xIDGrupo = cGrupo.ItemData(cGrupo.ListIndex)
    
    If MsgBox("¿Confirma almacenar los datos ingresados?", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    
    On Error GoTo errBT
    FechaDelServidor
    cBase.BeginTrans
    On Error GoTo errRB
    
    IdBML = 0
    cons = "Select * From BMLocal " _
            & " Where BMLLocal = " & cLocal.ItemData(cLocal.ListIndex) _
            & " And BMLArea = '" & xIDQue & "'" _
            & " And BMLCodigo = " & xIDGrupo
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        IdBML = rsAux!BMLId
        rsAux.Edit
        If Trim(tComentario.Text) = "" Then rsAux!BMLComentario = Null Else rsAux!BMLComentario = Trim(tComentario.Text)
        rsAux!BMLFModificacion = Format(gFechaServidor, sqlFormatoFH)
        rsAux!BMLUsuario = tUsuario.Tag
        rsAux.Update
        rsAux.Close
    Else
        rsAux.AddNew
        rsAux!BMLLocal = cLocal.ItemData(cLocal.ListIndex)
        rsAux!BMLArea = xIDQue
        rsAux!BMLCodigo = xIDGrupo
        If Trim(tComentario.Text) = "" Then rsAux!BMLComentario = Null Else rsAux!BMLComentario = Trim(tComentario.Text)
        rsAux!BMLFModificacion = Format(gFechaServidor, sqlFormatoFH)
        rsAux!BMLUsuario = tUsuario.Tag
        rsAux.Update
        rsAux.Close
        
        cons = "Select MAX(BMLId) From BMLocal"
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        IdBML = rsAux(0)
        rsAux.Close
    End If
    
    cons = "Select * From BMRenglon Where BMRIdBML = " & IdBML
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        rsAux.Delete: rsAux.MoveNext
    Loop
    rsAux.Close
    
'    cons = "Select * From BMRenglon Where BMRIdBML = " & IdBML
'    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        
'    With vsConsulta
'        For I = 1 To .Rows - 1
'            rsAux.AddNew
'            rsAux!BMRIdBML = IdBML
'            rsAux!BMRArticulo = CLng(.Cell(flexcpData, I, 1))
'            rsAux!BMREstado = CLng(.Cell(flexcpData, I, 2))
'            rsAux!BMRCantidad = CLng(.Cell(flexcpText, I, 1))
'            rsAux.Update
'        Next
'    End With
    
    With vsConsulta
        For I = 1 To .Rows - 1
            
            'Grabo en la tabla del local.
            cons = "Select * From BMRenglon Where BMRIdBML = " & IdBML _
                & " And BMRArticulo = " & CLng(.Cell(flexcpData, I, 1)) & " And BMREstado = " & CLng(.Cell(flexcpData, I, 2))
            
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            If rsAux.EOF Then
                rsAux.AddNew
                rsAux!BMRIdBML = IdBML
                rsAux!BMRArticulo = CLng(.Cell(flexcpData, I, 1))
                rsAux!BMREstado = CLng(.Cell(flexcpData, I, 2))
                rsAux!BMRCantidad = CLng(.Cell(flexcpText, I, 1))
                rsAux.Update
           Else
                rsAux.Edit
                'rsAux!BMRCantidad = CLng(.Cell(flexcpText, I, 1))
                rsAux!BMRCantidad = rsAux!BMRCantidad + CLng(.Cell(flexcpText, I, 1))
                rsAux.Update
            End If
            
 '           'If CLng(.Cell(flexcpText, I, 1)) = 0 Then rsAux.Delete
            rsAux.Close
            
        Next I
    End With
    
    cBase.CommitTrans
    
    Screen.MousePointer = 0
    
    Me.Refresh
    If MsgBox("Desea imprimir los datos ingresados?", vbQuestion + vbYesNo, "Imprimir") = vbYes Then
    
        If chVista.Value = 1 Then
            AccionImprimir False
            vsListado.ZOrder 0
        Else
            chVista.Value = 1
        End If
    
        frmSetup.pControl = vsListado
        frmSetup.Show vbModal, Me
        Me.Refresh
        If frmSetup.pOK Then vsListado.PrintDoc , frmSetup.pPaginaD, frmSetup.pPaginaH
    End If
    DeshabilitoIngreso
    bNuevo.Enabled = True: bGrabar.Enabled = False: bCancelar.Enabled = False
    cLocal.Enabled = True: cQue.Enabled = True: cGrupo.Enabled = True
    Screen.MousePointer = 0
    Exit Sub
    
errPru:
    MsgBox "Error en : " & sErr, vbExclamation, "ATENCIÓN"
    Screen.MousePointer = 0
    Exit Sub
errBT:
    clsGeneral.OcurrioError "Ocurrio un error al iniciar la transacción.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
errRB:
    Resume ErrResumo
ErrResumo:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrio un error al intentar grabar los datos, reintente.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
End Sub
Private Sub DeshabilitoIngreso()

'    cLocal.ListIndex = -1: cLocal.Enabled = False: cLocal.BackColor = Inactivo
'    tArea.Text = "": cQue.Enabled = False: cQue.BackColor = Inactivo
'    tCodigo.Text = "": cGrupo.Enabled = False: cGrupo.BackColor = Inactivo
    tArticulo.Text = "": tArticulo.Enabled = False: tArticulo.BackColor = Inactivo
    tCantidad.Text = vbNullString: tCantidad.Enabled = False: tCantidad.BackColor = Inactivo
    cEstado.Enabled = False: cEstado.BackColor = Inactivo: cEstado.ListIndex = -1
    tComentario.BackColor = Inactivo: tComentario.Enabled = False: tComentario.Text = vbNullString
    vsConsulta.Rows = 1
    tUsuario.Enabled = False: tUsuario.BackColor = Inactivo: tUsuario.Text = vbNullString: tUsuario.Tag = vbNullString
End Sub

Private Sub HabilitoIngreso()
    
    cLocal.Enabled = True: cLocal.BackColor = Obligatorio
    cQue.Enabled = True: cQue.BackColor = Obligatorio
    tArticulo.Enabled = True: tArticulo.BackColor = Obligatorio
    cGrupo.Enabled = True: cGrupo.BackColor = Obligatorio
    tCantidad.Text = vbNullString: tCantidad.Enabled = True: tCantidad.BackColor = Obligatorio
    cEstado.Enabled = True: cEstado.BackColor = Obligatorio
    tComentario.BackColor = Blanco: tComentario.Enabled = True
    tUsuario.Enabled = True: tUsuario.BackColor = Obligatorio: tUsuario.Text = vbNullString: tUsuario.Tag = vbNullString
    
End Sub

Private Sub AccionCancelar()
    If sNuevo Then
        bNuevo.Enabled = True: bGrabar.Enabled = False: bCancelar.Enabled = False: bModificar.Enabled = False
        labUsuario.Caption = ""
    Else
        bNuevo.Enabled = True: bGrabar.Enabled = False: bCancelar.Enabled = False: bModificar.Enabled = True
        cLocal.Enabled = True: cQue.Enabled = True: cGrupo.Enabled = True
    End If
    sNuevo = False
    DeshabilitoIngreso
End Sub

Private Sub AccionNuevo()
    sNuevo = True
    chVista.Value = 0
    bNuevo.Enabled = False: bModificar.Enabled = False: bGrabar.Enabled = True: bCancelar.Enabled = True
    HabilitoIngreso
    cLocal.Text = "": tComentario.Text = vbNullString: tComentario.Text = "": labUsuario.Caption = "": cQue.ListIndex = -1: cGrupo.ListIndex = -1
    Foco tArticulo  'Chinura para que agarre la primera vez
    cLocal.SetFocus
End Sub
Private Sub AccionModificar()
    CargoBalance
    chVista.Value = 0
    bNuevo.Enabled = False: bGrabar.Enabled = True: bCancelar.Enabled = True: bModificar.Enabled = False
    HabilitoIngreso
    cLocal.Enabled = False: cQue.Enabled = False: cGrupo.Enabled = False
    Foco tArticulo  'Chinura para que agarre la primera vez
End Sub
Private Sub InsertoRenglon()
On Error GoTo ErrControl
Dim aValor As Long

    If Val(tArticulo.Tag) = 0 Then MsgBox "No hay seleccionado un artículo.", vbExclamation, "ATENCIÓN": Foco tArticulo: Exit Sub
    If Not IsNumeric(tCantidad.Text) Then MsgBox "La cantidad ingresada no es correcta.", vbInformation, "ATENCIÓN": Foco tCantidad: Exit Sub
    If Val(tCantidad.Text) = 0 Then MsgBox "La cantidad ingresada no es correcta.", vbInformation, "ATENCIÓN": Foco tCantidad: Exit Sub
    If cEstado.ListIndex = -1 Then MsgBox "No hay seleccionado un estado.", vbInformation, "ATENCIÓN": Foco cEstado: Exit Sub
    
    'Busco en la grilla si ya se ingreso este artículo para ese estado.
    With vsConsulta
        For I = 1 To .Rows - 1
            If CLng(.Cell(flexcpData, I, 1)) = CLng(tArticulo.Tag) And CLng(.Cell(flexcpData, I, 2)) = cEstado.ItemData(cEstado.ListIndex) Then
            
                Dim pstrMSG As String, pintRES As Integer
                
                pstrMSG = "El artículo ya está ingresado." & vbCrLf & _
                          "¿Qué es lo que usted quiere hacer?  " & vbCrLf & vbCrLf & _
                          "SI: sustituye los «[QV]» ya ingresados por los «[QN]» de ahora." & vbCrLf & _
                          "NO: suma los «[QN]» ingresados a los «[QV]» ya existentes." & vbCrLf & _
                          "Cancelar: deja las cosas como están."
                          
                pstrMSG = Replace(pstrMSG, "[QV]", .Cell(flexcpText, I, 1))
                pstrMSG = Replace(pstrMSG, "[QN]", Val(tCantidad.Text))
                
                pintRES = MsgBox(pstrMSG, vbQuestion + vbYesNoCancel + vbDefaultButton3, "Artículo Ingresado")
                Select Case pintRES
                    Case vbCancel: Foco tArticulo: Exit Sub
                    
                    Case vbYes: .Cell(flexcpText, I, 1) = Val(tCantidad.Text)
                    Case vbNo: .Cell(flexcpText, I, 1) = .Cell(flexcpValue, I, 1) + Val(tCantidad.Text)
                End Select
                tArticulo.Text = "": tCantidad.Text = "": cEstado.Text = "": Foco tArticulo: Exit Sub
                
                
                'If MsgBox("El artículo ya está ingresado" & vbCrLf & vbCrLf & _
                         "¿Desea modificar la cantidad ingresada?", vbQuestion + vbYesNo, "Artículo Ingresado") = vbNo Then
                '         tArticulo.Text = "": tCantidad.Text = "": cEstado.Text = ""
                '         Foco tArticulo: Exit Sub
                'End If

                '.Cell(flexcpText, I, 1) = .Cell(flexcpValue, I, 1) + Val(tCantidad.Text)
                'tArticulo.Text = "": tCantidad.Text = "": cEstado.Text = "": Foco tArticulo: Exit Sub
            End If
        Next I
    End With
    
    Screen.MousePointer = 11
    With vsConsulta
        .AddItem ""
        
        .Cell(flexcpData, .Rows - 1, 1) = tArticulo.Tag
        .Cell(flexcpData, .Rows - 1, 2) = cEstado.ItemData(cEstado.ListIndex)
        
        .Cell(flexcpText, .Rows - 1, 1) = Val(tCantidad.Text)
        .Cell(flexcpText, .Rows - 1, 2) = Trim(tArticulo.Text)
        .Cell(flexcpText, .Rows - 1, 3) = Trim(cEstado.Text)
        .Cell(flexcpData, .Rows - 1, 3) = 0     'Me digo que lo ingreso ahora.
        .Select .Rows - 1, 0, .Rows - 1, .Cols - 1
        aValor = .CellTop
    End With
    tCantidad.Text = "": cEstado.Text = "": Foco tArticulo
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrControl:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrio un error al intentar insertar el renglón en la lista.", Err.Description
End Sub


Private Function BuscoArticuloPorCodigo(CodArticulo As String) As Boolean
'Atención el mapeo de error lo hago antes de entrar al procedimiento
Dim mBCode As String

    Screen.MousePointer = 11
    mBCode = Replace(CodArticulo, "'", "''")
    
    cons = "Select * From Articulo " & _
           " Where ArtCodigo = " & CodArticulo & _
            " OR ArtID IN (" & _
                    " Select Distinct(ACBArticulo) from ArticuloCodigoBarras " & _
                    " Where (ACBCodigo = '" & mBCode & "' And ACBLargo=0) Or ('" & mBCode & "' Like ACBCodigo And ACBLargo = " & Len(mBCode) & ")" & _
                ")"

    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If rsAux.EOF Then
        rsAux.Close
        tArticulo.Tag = "0"
        'MsgBox "No existe un artículo que posea ese código.", vbExclamation, "ATENCIÓN"
    Else
        tArticulo.Text = Format(rsAux!ArtCodigo, "(#,000,000)") & " " & Trim(rsAux!ArtNombre)
        tArticulo.Tag = rsAux!ArtID
        rsAux.Close
    End If
    Screen.MousePointer = 0

End Function

Private Sub BuscoArticuloPorNombre(NomArticulo As String)
'Atención el mapeo de error lo hago antes de entrar al procedimiento
Dim Resultado As Long
Dim objAyuda As clsListadeAyuda

Dim mBCode As String, mSQLBCode As String

    mBCode = Replace(NomArticulo, "'", "''")

    'mSQLBCode = " OR ArtID IN (" & _
                    " Select Distinct(ArtId) from Articulo " & _
                    " Left Outer Join ArticuloCodigoBarras ON ArtID = ACBArticulo " & _
                    " Where (ACBCodigo = '" & mBCode & "' And ACBLargo=0) Or ('" & mBCode & "' Like ACBCodigo And ACBLargo = " & Len(mBCode) & ")" & _
                ")"
                
    mSQLBCode = " OR ArtID IN (" & _
                    " Select Distinct(ACBArticulo) from ArticuloCodigoBarras " & _
                    " Where (ACBCodigo = '" & mBCode & "' And ACBLargo=0) Or ('" & mBCode & "' Like ACBCodigo And ACBLargo = " & Len(mBCode) & ")" & _
                ")"
                
    
    Screen.MousePointer = 11
    cons = "Select Count(*) From Articulo " _
        & " Where ArtNombre LIKE '" & Replace(NomArticulo, " ", "%") & "%'" & _
        mSQLBCode
        
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not IsNull(rsAux(0)) Then
        Resultado = rsAux(0)
    Else
        Resultado = 0
    End If
    rsAux.Close
    
    If Resultado = 0 Then
        MsgBox "No se encontraron artículos que coincidan con el dato ingresados.", vbInformation, "ATENCIÓN"
    Else
        If Resultado = 1 Then
            cons = "Select * From Articulo " _
                & " Where ArtNombre LIKE '" & Replace(NomArticulo, " ", "%") & "%'" & _
                mSQLBCode
                
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurReadOnly)
            Resultado = rsAux!ArtCodigo
            rsAux.Close
        Else
            cons = "Select ArtCodigo as Codigo, ArtNombre as Nombre from Articulo" _
                & " Where ArtNombre LIKE '" & Replace(NomArticulo, " ", "%") & "%'" & _
                mSQLBCode & _
                " Order By ArtNombre"
            
            Set objAyuda = New clsListadeAyuda
            If objAyuda.ActivarAyuda(cBase, cons, 4500, 0, "Ayuda de Artículos") > 0 Then
                Screen.MousePointer = 11
                Resultado = objAyuda.RetornoDatoSeleccionado(0)
            Else
                Resultado = 0
            End If
            Set objAyuda = Nothing
        End If
    End If
    If Resultado > 0 Then BuscoArticuloPorCodigo CStr(Resultado)
    Screen.MousePointer = 0
       
End Sub

Private Sub CargoBalance()
On Error GoTo ErrCB
Dim aValor As Long

    'Es consulta
    If bGrabar.Enabled = False Then tComentario.Text = "": labUsuario.Caption = "": vsConsulta.Rows = 1
    
    If cLocal.ListIndex = -1 Or cQue.ListIndex = -1 Or cGrupo.ListIndex = -1 Then Exit Sub
    
    Dim xIDQue As Long, xIDGrupo As Long
    
    xIDQue = cQue.ItemData(cQue.ListIndex)
    xIDGrupo = cGrupo.ItemData(cGrupo.ListIndex)
    
    cons = "Select * From BMLocal, BMRenglon, Articulo, EstadoMercaderia " _
        & " Where BMLLocal = " & cLocal.ItemData(cLocal.ListIndex) _
        & " And BMLArea = '" & xIDQue & "'" _
        & " And BMLCodigo = " & xIDGrupo _
        & " And BMLId = BMRIdBML And BMRArticulo = ArtID And BMREstado = ESMCodigo"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsAux.EOF Then
        If sNuevo Then
            If MsgBox("Se encontraron datos para los datos ingresados. ¿Desea proceder a cargar los datos del mismo?", vbQuestion + vbYesNo, "Cargar datos") = vbNo Then rsAux.Close: Exit Sub
            vsConsulta.Rows = 1
            'Ahora simulo que esta modificando
            sNuevo = False: cLocal.Enabled = False: cQue.Enabled = False: cGrupo.Enabled = False
        Else
            bModificar.Enabled = True
        End If
        
        If Not IsNull(rsAux!BMLComentario) Then tComentario.Text = Trim(rsAux!BMLComentario) Else tComentario.Text = ""
        labUsuario.Caption = z_BuscoUsuario(rsAux!BMLUsuario, True)
        
    End If
    Screen.MousePointer = 11
    Do While Not rsAux.EOF
        With vsConsulta
            .AddItem ""
'            If rsAux!BMRCantidad < 0 Then
'                .AddItem "Baja"
'                .Cell(flexcpData, .Rows - 1, 0) = 0
'                .Cell(flexcpPicture, .Rows - 1, 0) = ImgList.ListImages("Baja").ExtractIcon
'            Else
'                .AddItem "Alta"
'                .Cell(flexcpData, .Rows - 1, 0) = 1
'                .Cell(flexcpPicture, .Rows - 1, 0) = ImgList.ListImages("Alta").ExtractIcon
'            End If
            aValor = rsAux!ArtID: .Cell(flexcpData, .Rows - 1, 1) = aValor
            aValor = rsAux!EsMCodigo: .Cell(flexcpData, .Rows - 1, 2) = aValor
            .Cell(flexcpText, .Rows - 1, 1) = rsAux!BMRCantidad
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!ArtCodigo, "(#,000,000)") & " " & Trim(rsAux!ArtNombre)
            .Cell(flexcpText, .Rows - 1, 3) = Trim(rsAux!ESMNombre)
            .Cell(flexcpData, .Rows - 1, 3) = 1     'Me digo que estaba en el stock.
        End With
        rsAux.MoveNext
    Loop
    rsAux.Close
    Screen.MousePointer = 0
    Exit Sub
ErrCB:
    clsGeneral.OcurrioError "Ocurrio un error al cargar los datos del balance.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function fnc_CargoCombosConteo()
    
    cQue.Clear: cGrupo.Clear
    
    If cLocal.ListIndex = -1 Then Exit Function
    
    Dim xIDLocal As Long, mSQL As String
    xIDLocal = cLocal.ItemData(cLocal.ListIndex)
    
    mSQL = "Select QCBCodigo, QCBNombre from QueContarBalance" & _
                " Where QCBLocal = " & xIDLocal
    CargoCombo mSQL, cQue

    mSQL = "Select GrBCodigo, GrBNombre from GrupoBalance" & _
                " Where GrBLocal = " & xIDLocal
    CargoCombo mSQL, cGrupo

End Function
