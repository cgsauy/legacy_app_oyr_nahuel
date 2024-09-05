VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmArrimar 
   BackColor       =   &H00000000&
   Caption         =   "EntregaArrimar"
   ClientHeight    =   10350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14055
   FillColor       =   &H00400000&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmArrimar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10350
   ScaleWidth      =   14055
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00CC9966&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   12255
      TabIndex        =   8
      Top             =   6720
      Width           =   12255
      Begin VB.TextBox tcant 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000006&
         Height          =   495
         Left            =   10800
         MaxLength       =   3
         TabIndex        =   9
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lArticulo 
         BackStyle       =   0  'Transparent
         Caption         =   "lava a mano juna gran siete"
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   240
         TabIndex        =   13
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label lNuevaCant 
         BackStyle       =   0  'Transparent
         Caption         =   "Real:"
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   9480
         TabIndex        =   12
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lbArrimado 
         BackStyle       =   0  'Transparent
         Caption         =   "Arrimado:"
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   5640
         TabIndex        =   11
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label ltengo 
         BackStyle       =   0  'Transparent
         Caption         =   "Necesito:"
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   3360
         TabIndex        =   10
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.PictureBox picMsg 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   14055
      TabIndex        =   6
      Top             =   9255
      Width           =   14055
      Begin VB.Label lMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "HOLA A TODOS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   855
         Left            =   2160
         TabIndex        =   7
         Top             =   60
         Width           =   2175
      End
      Begin VB.Image imgIcon 
         Height          =   720
         Left            =   1080
         Picture         =   "frmArrimar.frx":058A
         Top             =   120
         Width           =   720
      End
      Begin VB.Shape shMsg 
         BackColor       =   &H00CC9966&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   960
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   13095
      End
   End
   Begin VB.Timer tAviso 
      Left            =   14160
      Top             =   1680
   End
   Begin VB.Timer TArr 
      Left            =   14160
      Top             =   2280
   End
   Begin VB.TextBox tcBarra 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   8520
      Width           =   3615
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsGridArt 
      Height          =   7695
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   13573
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   0
      ForeColor       =   -2147483634
      BackColorFixed  =   -2147483630
      ForeColorFixed  =   -2147483639
      BackColorSel    =   -2147483630
      ForeColorSel    =   12648447
      BackColorBkg    =   -2147483637
      BackColorAlternate=   0
      GridColor       =   -2147483635
      GridColorFixed  =   -2147483635
      TreeColor       =   -2147483632
      FloodColor      =   0
      SheetBorder     =   0
      FocusRect       =   0
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrimar.frx":182B4
            Key             =   "exc"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrimar.frx":2FFEE
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrimar.frx":47D28
            Key             =   "inf"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrimar.frx":5FA62
            Key             =   "car"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrimar.frx":7779C
            Key             =   "sto"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   240
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrimar.frx":7806B
            Key             =   "exc"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrimar.frx":8FDA5
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrimar.frx":A7ADF
            Key             =   "inf"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrimar.frx":BF819
            Key             =   "sto"
         EndProperty
      EndProperty
   End
   Begin VB.Label lPrendido 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Apagado (xx para encender)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   10080
      TabIndex        =   14
      Top             =   180
      Width           =   3855
   End
   Begin VB.Label llist 
      BackStyle       =   0  'Transparent
      Caption         =   "Artículos para arrimar"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   105
      Width           =   12735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Editar:  A + Codigo de Barra"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   8880
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sumar: Cant. + (*) + Codigo de Barra"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   8520
      Width           =   5175
   End
   Begin VB.Label lcant 
      BackColor       =   &H00000080&
      Caption         =   " Cantidad de mas : "
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9000
      TabIndex        =   2
      Top             =   8640
      Width           =   3975
   End
   Begin VB.Shape shlist 
      BackColor       =   &H00CEBAB3&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00A88D7B&
      BorderWidth     =   2
      Height          =   495
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   12975
   End
   Begin VB.Menu MnuGrilla 
      Caption         =   "MnuGrilla"
      Visible         =   0   'False
      Begin VB.Menu MnuAnular 
         Caption         =   "Anular un documento"
      End
      Begin VB.Menu MnuClienteSeFue 
         Caption         =   "Anular documento el Cliente Se Fue"
      End
      Begin VB.Menu MnuLineResta 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRestoDocumento 
         Caption         =   "Restar artículos a un único documento"
      End
      Begin VB.Menu MnuQuitarDeDocumento 
         Caption         =   "Restar este artículo a todos los documentos"
      End
   End
End
Attribute VB_Name = "frmArrimar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'16/7/207   CargoArtículos para los contados y créditos puse union para identificar los artículos que son con detalles.

Option Explicit

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private sWav As String

Public Enum EstadoR 'estado de entrega
    SinEntregar = 1
    Arrimado = 2
    Entregado = 3
    ClienteSeFue = 4
    EraParaEnvio = 5
    OtroLocal = 6
End Enum

Private Sub loc_Help()
On Error GoTo errHelp
    Screen.MousePointer = 11
    
    Dim aFile As String
    Cons = "Select * from Aplicacion Where AplNombre = '" & App.Title & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux!AplHelp) Then aFile = Trim(RsAux!AplHelp)
    RsAux.Close
    If aFile <> "" Then EjecutarApp aFile
    Screen.MousePointer = 0
    Exit Sub
errHelp:
    clsGeneral.OcurrioError "Error al activar el archivo de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub loc_MsgPrendidoApagado()
    With lPrendido
        If paArrimar = 1 Then
            .Caption = "Encendido (xx para apagar)"
            .ForeColor = &H800000
            TArr.Interval = 10
            TArr.Enabled = True
        Else
            .Caption = "Apagado (xx para encender)"
            .ForeColor = &H40C0&
        End If
    End With
End Sub

Private Sub loc_PrendoApago()
On Error GoTo errPA
    
    Cons = "Select * from Parametro Where ParNombre = 'dep_Estado_Arrimar_" & paCodigoDeSucursal & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.AddNew
        RsAux("ParNombre") = "dep_Estado_Arrimar_" & paCodigoDeSucursal
    Else
        RsAux.Edit
    End If
    RsAux("ParValor") = IIf(paArrimar = 1, 0, 1)
    RsAux.Update
    RsAux.Close
    paArrimar = IIf(paArrimar = 1, 0, 1)
    
    If vsGridArt.Rows > 0 And paArrimar = 0 Then
        MsgBox "Atención quedan artículos pendientes de arrimar TIENE QUE ARRIMARLOS para que aparezcan en el formulario de entrega los documentos.", vbExclamation, "ATENCIÓN"
    End If
    loc_MsgPrendidoApagado
    Exit Sub
errPA:
    clsGeneral.OcurrioError "Error al cambiar el estado", Err.Description, "Prender/apagar"
End Sub

Private Sub loc_OcultoMsg()
    
    loc_ControlEdit False
    tcBarra.Enabled = True
    lcant.Visible = False
    lcant.Caption = "Cantidad de más :"
    OcultarMsg False
    tcBarra.SetFocus
    
End Sub

Private Sub loc_ControlEdit(estado As Boolean)
On Error Resume Next
    
    picEdit.Visible = estado
    
    If Not estado Then
        tcant.Text = ""
        ltengo.Caption = "Necesito:"
        lbArrimado.Caption = "Arrimado:"
        lbArrimado.Tag = ""
    End If
    
End Sub

Private Sub loc_ShowMsg(ByVal sTxt As String, ByVal iEstado As Integer)
    
    Select Case iEstado
        Case 1
            imgIcon.Picture = imlIcons.ListImages("exc").Picture
            loc_SetSonido paSonidoMal
        Case 2
            imgIcon.Picture = imlIcons.ListImages("inf").Picture
            loc_SetSonido paSonidoOK
        Case 3
            imgIcon.Picture = imlIcons.ListImages("sto").Picture
            loc_SetSonido paSonidoMal
    End Select
    
    lMsg.Caption = sTxt
    
    lMsg.Visible = True
    imgIcon.Visible = True
    shMsg.Visible = True
    
    tAviso.Interval = 4000
    tAviso.Enabled = True

End Sub

Private Function fnc_GetIDArt(ByVal sCodBar As String, ByRef iRetQ As Integer, Optional ByRef bEsEspecifico As Boolean = False) As Long
Dim sQy As String
Dim RsAux As rdoResultset
On Error GoTo errGI
    
    iRetQ = 1
    bEsEspecifico = False
    sCodBar = Replace(sCodBar, "'", "''")
    sQy = " Select Distinct(ArtId), ACBCantidad from Articulo Left Outer Join ArticuloCodigoBarras ON ArtID = ACBArticulo " _
        & "Where (ACBCodigo = '" & sCodBar & "' And ACBLargo=0) Or ('" & sCodBar & "' Like ACBCodigo And ACBLargo = " & Len(sCodBar) & ")"

    If IsNumeric(sCodBar) Then sQy = sQy & " Or ArtCodigo = " & sCodBar
    
    Set RsAux = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        fnc_GetIDArt = RsAux("ArtId")
        If Not IsNull(RsAux("ACBCantidad")) Then iRetQ = RsAux("ACBCantidad")
    End If
    RsAux.Close
    
    'Si no cargue entonces voy al artículo específico
    If fnc_GetIDArt = 0 Then
        sQy = "Select AEsID From (EntregaAuxiliar Inner Join ArticuloEspecifico ON AEsArticulo = EAuArticulo And EAuTipo = AEsTipoDocumento And AEsDocumento = EAuDocumento) " & _
                " Where AEsID = " & sCodBar & " And EAuEstado = 1"
        Set RsAux = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            bEsEspecifico = True
            fnc_GetIDArt = RsAux(0)
        End If
        RsAux.Close
    End If
    
Exit Function
errGI:
    clsGeneral.OcurrioError "Error al buscar el ID del artículo.", Err.Description, "Buscar artículo"
End Function
Private Sub loc_SetSonido(ByVal sFile As String)
    On Error Resume Next
    If sFile = "" Then Exit Sub
    Dim Result As Long
    Result = sndPlaySound(sWav & sFile, 1)
End Sub

Private Sub Form_Load()
On Error Resume Next

    picMsg.BackColor = vbBlack
    
    ChDir App.Path
    ChDir ("..")
    If Dir(CurDir & "\Sonidos", vbDirectory) <> "" Then sWav = CurDir & "\Sonidos\"

    loc_MsgPrendidoApagado
    ObtengoSeteoForm Me, 100, 100
    vsGridArt.BackColorBkg = vbBlack
    
    With vsGridArt
        .ColWidth(0) = 1000
        .Cols = 2
        .ColAlignment(1) = flexAlignLeftCenter
    End With
    
    loc_ControlEdit False
'    CargoArticulos
    lcant.Visible = False
    
    OcultarMsg False

    TArr.Interval = 10
    TArr.Enabled = True
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    
    tcBarra.Top = picMsg.Top - tcBarra.Height - 180
    Label1.Top = tcBarra.Top
    Label3.Top = Label1.Top + 360
    
    vsGridArt.Move vsGridArt.Left, vsGridArt.Top, Me.ScaleWidth - (vsGridArt.Left * 2), tcBarra.Top - vsGridArt.Top - 120
    shlist.Move shlist.Left, shlist.Top, (Me.ScaleWidth - (shlist.Left * 2)), vsGridArt.Top - shlist.Top - 120
    llist.Move llist.Left, llist.Top, shlist.Width, vsGridArt.Top - shlist.Top - 120
    
    shMsg.Move vsGridArt.Left, 0, vsGridArt.Width
    imgIcon.Left = shMsg.Left + 120
    lMsg.Move shMsg.Left + 1200, 60, shMsg.Width - 1440
    
    picEdit.Move shMsg.Left, picMsg.Top - 60, shMsg.Width
    tcant.Left = picEdit.Width - tcant.Width - 240
    lNuevaCant.Left = tcant.Left - lNuevaCant.Width - 120
    lbArrimado.Left = lNuevaCant.Left - lbArrimado.Width - 120
    ltengo.Left = lbArrimado.Left - ltengo.Left - 120
    lArticulo.Width = ltengo.Left - 120 - lArticulo.Left
        
    lPrendido.Left = (shlist.Left + shlist.Width) - lPrendido.Width
    lcant.Move vsGridArt.Width - lcant.Width, tcBarra.Top + 120
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    cBase.Close
    eBase.Close
    GuardoSeteoForm Me
End Sub

Private Sub MnuAnular_Click()
On Error Resume Next
    CambioEstadoArt 4, vsGridArt.Cell(flexcpData, vsGridArt.Row, 0), Val(vsGridArt.Cell(flexcpData, vsGridArt.Row, 1))
End Sub

Private Sub MnuClienteSeFue_Click()
On Error Resume Next
    CambioEstadoArt 5, vsGridArt.Cell(flexcpData, vsGridArt.Row, 0), Val(vsGridArt.Cell(flexcpData, vsGridArt.Row, 1))
End Sub

Private Sub MnuQuitarDeDocumento_Click()
On Error GoTo errQD
    If MsgBox("¿Confirma anular la engrega de " & vsGridArt.Cell(flexcpText, vsGridArt.Row, 1) & "?" & vbCrLf & vbCrLf & _
            "Todos los documentos que posean este artículo y no tengan otro por entregar serán anulados.", vbQuestion + vbYesNo, "Quitar un artículo de entrega") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    TArr.Enabled = False
    Cons = "Select * from EntregaAuxiliar Where EAuArticulo = " & vsGridArt.Cell(flexcpData, vsGridArt.Row, 0) & " And EAuEstado = 1 And EAuTipo <> 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        RsAux.Edit
        RsAux("EAuEstado") = 4
        RsAux.Update
        RsAux.MoveNext
    Loop
    RsAux.Close
    TArr.Interval = 10
    TArr.Enabled = True
    Screen.MousePointer = 0
    Exit Sub
errQD:
    clsGeneral.OcurrioError "Error al quitar los artículos.", Err.Description, "Quitar artículos a documentos."
    TArr.Interval = 10
    TArr.Enabled = True
    Screen.MousePointer = 0
End Sub

Private Sub MnuRestoDocumento_Click()
Dim iDoc As Long
Dim iQ As Integer
On Error GoTo errRD
    Cons = "Select EAuDocumento, IsNull(DocSerie, 'Rem') Serie, IsNull(DocNumero, RemCodigo) Numero, EAuFechaHora as Ingresó" _
        & " From ((EntregaAuxiliar LEFT Join Documento ON EAuDocumento = DocCodigo And EAuTipo <> 6)" _
        & " Left Join Remito ON EAuDocumento = RemCodigo And EAuTipo = 6)" _
        & " Where EAuEstado = 1 And EAuArticulo = " & vsGridArt.Cell(flexcpData, vsGridArt.Row, 0) & "Order by EAuFechaHora"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
       RsAux.Close
       MsgBox "El artículo no está más pendiente de arrimar, verifique.", vbExclamation, "Atención"
       Exit Sub
    Else
        RsAux.MoveNext
        If RsAux.EOF Then RsAux.MoveFirst: iDoc = RsAux(0)
        RsAux.Close
    End If
    If iDoc = 0 Then iDoc = fnc_GetDocumentoArticuloHelp(Cons)
    If iDoc = 0 Then Exit Sub
        
    Cons = "Select * from EntregaAuxiliar Where EAuDocumento = " & iDoc & " And EAuArticulo = " & vsGridArt.Cell(flexcpData, vsGridArt.Row, 0) & " And EAuEstado = 1 And EAuTipo <> 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    iQ = RsAux("EAuSinArrimar")
    
    Cons = InputBox("Ingrese la cantidad de artículos a restar (se presenta la cantidad que falta por arrimar)", "Restar artículos", iQ)
    
    If Not IsNumeric(Cons) Or Val(Cons) < 1 Then
        RsAux.Close
        MsgBox "No ingresó un valor correcto.", vbExclamation, "Atención"
        Exit Sub
    Else
        If CInt(Cons) > iQ Then
            RsAux.Close
            MsgBox "La cantidad ingresada supera lo pendiente de arrimar.", vbExclamation, "Atención"
            Exit Sub
        End If
    End If
    
    iQ = CInt(Cons)
    
    If MsgBox("¿Confirma restarle al documento la cantidad a entregar de  " & Cons & " '" & vsGridArt.Cell(flexcpText, vsGridArt.Row, 1) & "'?" & vbCrLf & vbCrLf _
            , vbQuestion + vbYesNo, "Quitar un artículo de entrega") = vbNo Then RsAux.Close: Exit Sub
    Screen.MousePointer = 11
    TArr.Enabled = False
    
    RsAux.Edit
    If RsAux("EAuCantidad") - iQ = 0 Then
        'Es todo
        RsAux("EAuEstado") = 4
    Else
        RsAux("EAuCantidad") = RsAux("EAuCantidad") - iQ
        RsAux("EAuSinArrimar") = RsAux("EAuSinArrimar") - iQ
        If RsAux("EAuSinArrimar") = 0 Then RsAux("EAuEstado") = 2
    End If
    RsAux.Update
    RsAux.Close
    TArr.Interval = 10
    TArr.Enabled = True
    Screen.MousePointer = 0
    Exit Sub
errRD:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al restar artículos.", Err.Description, "Restar artículos"
    TArr.Interval = 10
    TArr.Enabled = True
End Sub

Private Sub TArr_Timer()
    TArr.Enabled = False
    TArr.Interval = 10000
    If paArrimar = 1 Then CargoArticulos
End Sub

Private Sub tAviso_Timer()
    
    tAviso.Enabled = False
    loc_OcultoMsg
    
End Sub

Private Sub OcultarMsg(ByVal estado As Boolean)
    lMsg.Visible = estado
    shMsg.Visible = estado
    imgIcon.Visible = estado
End Sub

Private Sub tcant_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        loc_OcultoMsg
        TArr.Enabled = True
    End If
End Sub

Private Sub tcant_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tcant.Text <> "" And IsNumeric(tcant.Text) Then
            loc_NuevaCantArr tcant.Text
            CargoArticulos
            If lcant.Visible Then tAviso.Interval = 2000 Else tAviso.Interval = 100
            tAviso.Enabled = True
            tcBarra.Text = ""
        End If
    End If
End Sub

Private Sub tcBarra_GotFocus()
On Error Resume Next
    With tcBarra
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tcBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then loc_Help
End Sub

Private Sub tcBarra_KeyPress(KeyAscii As Integer)
On Error Resume Next
    
    If KeyAscii = vbKeyReturn And Trim(tcBarra.Text) <> "" Then
        TArr.Enabled = False
        tAviso.Enabled = False
        loc_OcultoMsg
    
        If LCase(Left(tcBarra.Text, 1)) = "a" And Len(tcBarra.Text) > 1 Then
            loc_Edit tcBarra.Text
            Exit Sub
        ElseIf LCase(Left(tcBarra.Text, 1)) = "s" And Len(tcBarra.Text) > 1 Then
            If Not IsNumeric(Mid(tcBarra.Text, 2)) Then
                loc_ShowMsg "Servicio incorrecto", 1
                Exit Sub
            End If
            loc_BuscoServicio Mid(tcBarra.Text, 2)
        ElseIf LCase(Left(tcBarra.Text, 1)) = "." And Len(tcBarra.Text) > 1 Then
            AgregarArticulo Mid(tcBarra.Text, 2)
            Exit Sub
        ElseIf Trim(LCase(tcBarra.Text)) = "xx" Then
            loc_PrendoApago
        Else
            loc_SumoArrimado tcBarra.Text
            CargoArticulos
            tAviso.Interval = 3000
            tAviso.Enabled = True
        End If
        tcBarra.Text = ""
    End If
End Sub
Private Function fnc_Qstock(ByVal idArt As Long) As Long
Dim mSQL As String
Dim rsA As rdoResultset
On Error GoTo errStock
    
    mSQL = " Select STLCantidad From StockLocal Where STLArticulo = " & idArt & " And STLEstado =" & paEstadoArticuloEntrega & _
           " And STLLocal =" & paCodigoDeSucursal
    Set rsA = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    If Not rsA.EOF Then fnc_Qstock = rsA("STLCantidad")
    rsA.Close
    Exit Function
errStock:
    fnc_Qstock = 0
    'clsGeneral.OcurrioError "Ocurrió un error al verificar la cantidad de stock.", Err.Description
End Function

Private Sub CargoArticulos()
Dim Cons As String
Dim RsAux As rdoResultset
Dim lAux As Long
On Error GoTo errCargar
    
    Dim bVacio As Boolean
    bVacio = (vsGridArt.Rows = 0)
    vsGridArt.Rows = 0
    
    Screen.MousePointer = 11
    Set RsAux = cBase.OpenResultset("exec prg_arrimar_cargaarticulos", rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        If RsAux("Cantidad") > 0 Then
            With vsGridArt
                .AddItem RsAux("Cantidad")
                .Cell(flexcpText, .Rows - 1, 1) = " " & Trim(RsAux("ArtNombre"))
                If RsAux("Tipo") = 1 Then
                    lAux = RsAux("ArtID")
                    .Cell(flexcpData, .Rows - 1, 0) = lAux
                    .Cell(flexcpData, .Rows - 1, 1) = -1
                    lAux = fnc_Qstock(RsAux("ArtID"))
                    If lAux > 0 Then
                        If RsAux("Cantidad") > 1 Then .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = &HC0FFC0
                    Else
                        .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = &H80FF&
                    End If
                Else
                     .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = &H80FFFF
                     .Cell(flexcpText, .Rows - 1, 1) = " (Servicio " & RsAux("Servicio") & ") " & Trim(RsAux("ArtNombre"))
                     lAux = RsAux("Documento")
                     .Cell(flexcpData, .Rows - 1, 0) = lAux
                     lAux = RsAux("TipoDoc")
                     .Cell(flexcpData, .Rows - 1, 1) = lAux
                End If
            End With
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    Screen.MousePointer = 0
    If bVacio And vsGridArt.Rows > 0 Then loc_SetSonido paSonidoTimbre
    On Error Resume Next
    tcBarra.SetFocus
    TArr.Enabled = True
    Exit Sub
errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los artículos.", Err.Description
End Sub

Private Sub loc_SumoArrimado(ByVal Texto As String)

On Error GoTo errSumo
Dim Cons As String
Dim RsAux As rdoResultset
Dim sCodBar As String
Dim cant As Integer, iQ As Integer
Dim idArt As Integer
Dim bMod As Boolean 'variable para saber si se modificó algún registro
Dim bEspecifico As Boolean

    If InStr(1, Texto, "*") > 0 Then
        If Not IsNumeric(Mid(Texto, 1, InStr(Texto, "*") - 1)) Then loc_ShowMsg "La cantidad no es numérica", 3: Exit Sub
        sCodBar = (Trim(Mid(Texto, InStr(Texto, "*") + 1, Len(Texto))))
        cant = CInt(Mid(Texto, 1, InStr(Texto, "*") - 1))
    Else
        cant = 1
        sCodBar = Texto
    End If
            
    idArt = fnc_GetIDArt(sCodBar, iQ, bEspecifico)
    
    If idArt = 0 Then loc_ShowMsg "Artículo inexistente.", 3: Exit Sub
    
    cant = iQ * cant
    
    cBase.BeginTrans
    On Error GoTo errRB

    'Aquí tengo que controlar que no haya pasado el código de barras del artículo y no del artículo específico.
    If Not bEspecifico Then
        Cons = "Select EauSinArrimar, EAuEstado " _
            & " From (EntregaAuxiliar Inner Join Articulo ON EAuArticulo = ArtId)" _
            & " Where EAuEstado = 1 And ArtID = " & idArt _
            & " And Not Exists(Select * From ArticuloEspecifico Where AEsArticulo = ArtID And EAuTipo = AEsTipoDocumento And AEsDocumento = EAuDocumento)"
    
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux.EOF Then
            'Me marco señal para no ingresar.
            Cons = ""
        End If
        RsAux.Close
    Else
        Cons = "OK"
    End If
    
    If Cons <> "" Then
        'Si es artículo común
        If Not bEspecifico Then
            Cons = "Select EauSinArrimar, EAuEstado from EntregaAuxiliar " & _
               " Where EAuArticulo = " & idArt & " And EAuEstado in (1,2)" & _
               " And Not Exists(Select * From ArticuloEspecifico Where AEsArticulo = EAuArticulo And EAuTipo = AEsTipoDocumento And AEsDocumento = EAuDocumento)" & _
               " Order By EauFechaHora ASC "
        Else
            Cons = "Select EAuArticulo, EauSinArrimar, EAuEstado From (EntregaAuxiliar Inner Join ArticuloEspecifico ON AEsArticulo = EAuArticulo And EAuTipo = AEsTipoDocumento And AEsDocumento = EAuDocumento) " & _
                " Where AEsID = " & idArt & " And EAuEstado = 1"
        End If
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        Cons = ""
        If Not RsAux.EOF Then
            Cons = "OK"
            Do While Not RsAux.EOF
                RsAux.Edit
            
                If RsAux("EAuSinArrimar") <= cant Then
                    cant = cant - RsAux("EAuSinArrimar")
                    RsAux("EAuSinArrimar") = 0
                    RsAux("EauEstado") = 2
                Else
                    RsAux("EAuSinArrimar") = RsAux("EAuSinArrimar") - cant
                    cant = 0
                End If
                RsAux.Update
                RsAux.MoveNext
                bMod = True 'variable para saber si se modificó algún registro
                If cant = 0 Then Exit Do
            Loop
            If cant > 0 Then lcant.Visible = True: lcant.Caption = "Cantidad de más:  " & cant
        End If
        RsAux.Close
        cBase.CommitTrans
    End If
    If Cons = "" Then loc_ShowMsg "El artículo ingresado no está para arrimar, VERIFIQUE", 1
    
Exit Sub
errRoll:
    cBase.RollbackTrans
    Exit Sub

errRB:
    Resume errRoll

errSumo:
    Exit Sub
    clsGeneral.OcurrioError "Ocurrió un error al cargar los artículos sin arrimar.", Err.Description
End Sub

Private Sub loc_Edit(ByVal Texto As String)
On Error GoTo errEdit
Dim sCodBar As String
Dim Cons As String
Dim RsAux As rdoResultset
Dim idArt As Integer

    sCodBar = Mid(Texto, 2)
    
    sCodBar = Replace(sCodBar, "'", "''")
    Cons = "Select Distinct(ArtId), ArtNombre from Articulo Left Outer Join ArticuloCodigoBarras ON ArtID = ACBArticulo " _
        & "Where (ACBCodigo = '" & sCodBar & "' And ACBLargo=0) Or ('" & sCodBar & "' Like ACBCodigo And ACBLargo = " & Len(sCodBar) & ")"
        
    If IsNumeric(sCodBar) Then
        Cons = Cons & " Or ArtCodigo = " & sCodBar
    End If
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        idArt = RsAux("ArtId")
        lArticulo.Caption = RsAux("ArtNombre")
        lArticulo.Tag = idArt
    End If
    RsAux.Close
    If idArt = 0 Then loc_ShowMsg "El artículo no existe", 3: Exit Sub
    
    Cons = " Select cantSA = IsNull(Sum(EauSinArrimar),0) , Cant = IsNull(Sum(EAuCAntidad), 0) from EntregaAuxiliar " _
            & " Where EAuArticulo =" & idArt & " And EAuEstado In (1,2) And EauLocal = " & paCodigoDeSucursal

    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
      ltengo.Caption = ltengo.Caption & " " & RsAux("cant")
      lbArrimado.Caption = "Arrimado: " & RsAux("Cant") - RsAux("cantSA")
      lbArrimado.Tag = RsAux("Cant") - RsAux("cantSA")
      idArt = RsAux("Cant")
    End If
    RsAux.Close
    
    If idArt = 0 Then
        loc_ShowMsg "El artículo no tiene nada para arrimar", 3
    Else
        tcBarra.Enabled = False
        loc_ControlEdit True
        Foco tcant
    End If

Exit Sub
errEdit:
    clsGeneral.OcurrioError "Ocurrió un error al cargar el artículo para editar.", Err.Description
End Sub

Private Sub loc_NuevaCantArr(ByVal cant As Integer)
On Error GoTo errNuevaCant
Dim UltFecha As Date
Dim sCons As String
Dim RsAux As rdoResultset
    
    If Val(lbArrimado.Tag) < cant Then loc_SumoArrimado ((cant - Val(lbArrimado.Tag)) & "*" & (Mid(tcBarra.Text, 2))): Exit Sub
    
    sCons = " Select EAuSinArrimar, EAuEstado, EAuCantidad from EntregaAuxiliar " _
          & "  where EAuArticulo = " & lArticulo.Tag & " and EAuLocal = " & paCodigoDeSucursal _
          & " and EAuEstado In (1,2)  order by EAuFechaHora desc"
          
    Set RsAux = cBase.OpenResultset(sCons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        RsAux.Edit
        RsAux("EAuSinArrimar") = (Val(lbArrimado.Tag) - cant) + RsAux("EAuSinArrimar")
        RsAux("EAuEstado") = 1
        If RsAux("EAuSinArrimar") = 0 Then RsAux("EAuEstado") = 2
        RsAux.Update
        RsAux.MoveNext
    End If
    RsAux.Close
    
Exit Sub
errNuevaCant:
    clsGeneral.OcurrioError "Ocurrió un error al guardar la cantidad arrimada.", Err.Description
End Sub

Private Sub loc_BuscoServicio(ByVal pCodigoSer As Long)
On Error GoTo errServicio
Dim sCons As String
Dim lDocumento As Long, iTipo As Byte
Dim RsAux As rdoResultset
    iTipo = 3
    sCons = " Select SerCodigo, serDocumento From Servicio Where SerCodigo =" & pCodigoSer
    Set RsAux = cBase.OpenResultset(sCons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        If Not IsNull(RsAux("SerDocumento")) Then
            lDocumento = RsAux("serDocumento")
            iTipo = 1
        Else
            iTipo = 0
            lDocumento = pCodigoSer
        End If
    End If
    RsAux.Close
    
    'si el itipo = 3 no hay servicio
    If iTipo = 3 Then loc_ShowMsg "Servicio inexistente", 1: Exit Sub
    
    'aca estoy grabando el servicio
    sCons = " select EAuEstado, EAuSinArrimar from EntregaAuxiliar where EAuDocumento =" & lDocumento & " and EAuEstado In (1,2)"
    If iTipo = 1 Then
        sCons = sCons & " and EAuTipo In (1,2)"
    Else
        sCons = sCons & " and EAuTipo = 0"
    End If
    Set RsAux = cBase.OpenResultset(sCons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        RsAux.Edit
        RsAux("EAuEstado") = 2
        RsAux("EAuSinArrimar") = 0
        RsAux.Update
    Else
        RsAux.Close
        sCons = " select EAuEstado, EAuSinArrimar from EntregaAuxiliar where EAuDocumento =" & pCodigoSer & " and EAuEstado In (1,2) and EAuTipo = 0"
        Set RsAux = cBase.OpenResultset(sCons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            RsAux.Edit
            RsAux("EAuEstado") = 2
            RsAux("EAuSinArrimar") = 0
            RsAux.Update
        Else
            loc_ShowMsg "No existe el servicio, VERIFIQUE", 1
        End If
    End If
    RsAux.Close
    CargoArticulos
    
Exit Sub
errServicio:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el servicio.", Err.Description
End Sub

Private Sub AgregarArticulo(ByVal sCodArt As String)
On Error GoTo errAA
Dim sCons As String
Dim RsAux As rdoResultset
Dim idArt As Long
   
    idArt = fnc_GetIDArt(sCodArt, 0)
    If idArt = 0 Then loc_ShowMsg "El artículo no existe", 1: Exit Sub
    
    sCons = "Select * from ArticuloStockEntrega where AseArticulo = " & idArt & "And AseLocal = " & paCodigoDeSucursal
    Set RsAux = cBase.OpenResultset(sCons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        RsAux.Edit
        If RsAux("AseCercano") Then
            RsAux("AseCercano") = False
            sCons = "LEJANO"
        Else
            RsAux("AseCercano") = True
            sCons = "CERCANO"
        End If
    Else
        sCons = "CERCANO"
        RsAux.AddNew
        RsAux("AseCercano") = True
        RsAux("AseLocal") = paCodigoDeSucursal
        RsAux("ASEStkEmergencia") = 0
        RsAux("AseArticulo") = idArt
    End If
    RsAux.Update
    RsAux.Close
    loc_ShowMsg "El artículo quedó como " & sCons, 2
    tcBarra.Text = ""
    CargoArticulos
    Exit Sub
errAA:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el Artículo.", Err.Description
End Sub

Private Sub vsGridArt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errMD
    If vsGridArt.Rows > 0 And Button = 2 Then
        If vsGridArt.Row >= 0 Then
            MnuRestoDocumento.Enabled = (Val(vsGridArt.Cell(flexcpData, vsGridArt.Row, 1)) = -1)
            MnuQuitarDeDocumento.Enabled = (Val(vsGridArt.Cell(flexcpData, vsGridArt.Row, 1)) = -1)
            TArr.Enabled = False
            PopupMenu MnuGrilla
            TArr.Enabled = True
        End If
    End If
errMD:
End Sub

Private Function fnc_GetDocumentoArticuloHelp(ByVal sQy As String) As Long
On Error GoTo errGD
    Dim oHelp As New clsListadeAyuda
    If oHelp.ActivarAyuda(cBase, sQy, 3500, 1, "Documentos a entregar") > 0 Then
        fnc_GetDocumentoArticuloHelp = oHelp.RetornoDatoSeleccionado(0)
    End If
    Set oHelp = Nothing
errGD:
End Function

Private Sub CambioEstadoArt(ByVal iEstado As Byte, iDoc As Long, iTipo As Integer)
On Error GoTo errCEArt
Dim Cons As String
Dim RsAux As rdoResultset
    
    If iTipo < 0 Then
        'busco todos los documentos que tengan este artículo y lo presento.
        Cons = "Select EAuDocumento, IsNull(DocSerie, 'Re') Serie, IsNull(DocNumero, RemCodigo) Numero, EAuFechaHora as Ingresó" _
            & " From ((EntregaAuxiliar LEFT Join Documento ON EAuDocumento = DocCodigo And EAuTipo <> 6)" _
            & " Left Join Remito ON EAuDocumento = RemCodigo And EAuTipo = 6)" _
            & " Where EAuEstado = 1 And EAuArticulo = " & iDoc & "Order by EAuFechaHora"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        iDoc = 0
        If RsAux.EOF Then
            RsAux.Close
            MsgBox "No hay un documento pendiente con ese artículo.", vbExclamation, "Atención"
            Exit Sub
        Else
            RsAux.MoveNext
            If RsAux.EOF Then RsAux.MoveFirst: iDoc = RsAux(0)
            RsAux.Close
        End If
        If iDoc = 0 Then
            iDoc = fnc_GetDocumentoArticuloHelp(Cons)
        End If
        
        If iDoc = 0 Then Exit Sub
        
        If MsgBox("¿Confirma cambiar el estado al documento?" & vbCrLf & vbCrLf & _
                "Al aceptar TODOS LOS ARTÍCULOS del documento desapareceran de la lista y el documento no será incluido en la lista de entrega final.", _
            vbQuestion + vbYesNo, "Anular o Cliente se fue") = vbNo Then Exit Sub
    
        Cons = "Select EauEstado from EntregaAuxiliar Where EAuDocumento = " & iDoc & " And EAuTipo In (1,2,6) And EAuEstado = 1"
        
    Else
        
        If MsgBox("¿Confirma cambiar el estado al servicio?" & vbCrLf & vbCrLf & _
                "Al aceptar el servicio no será incluido en la lista de entrega final.", _
            vbQuestion + vbYesNo, "Anular o Cliente se fue") = vbNo Then Exit Sub
        
        Cons = "Select EauEstado from EntregaAuxiliar Where EauDocumento = " & iDoc & " and EauTipo = " & iTipo
    End If
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        Do While Not RsAux.EOF
            RsAux.Edit
            RsAux("EAuEstado") = iEstado
            RsAux.Update
            RsAux.MoveNext
        Loop
    End If
    RsAux.Close
    lMsg.Caption = "Cambio de estado correcto": tAviso.Enabled = True: tAviso.Interval = 3000: loc_SetSonido paSonidoOK
    OcultarMsg True
    CargoArticulos
    
Exit Sub
errCEArt:

End Sub


