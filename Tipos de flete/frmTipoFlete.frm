VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTipoFlete 
   Caption         =   "Tipos de Flete"
   ClientHeight    =   3675
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7080
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTipoFlete.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   7080
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "eliminar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancelar"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar staMensaje 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   3420
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsAgenda 
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   1380
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2566
      _ConvInfo       =   1
      Appearance      =   1
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
      BackColor       =   -2147483633
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483633
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
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
      Editable        =   -1  'True
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
   End
   Begin VB.TextBox tDescripcion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   25
      TabIndex        =   2
      Top             =   600
      Width           =   3255
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6420
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoFlete.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoFlete.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoFlete.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoFlete.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoFlete.frx":088A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Agenda de reparto"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2355
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Descripción:"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   600
      Width           =   1035
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuOpNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu MnuOpModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuOpEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
         Visible         =   0   'False
      End
      Begin VB.Menu MnuOpLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpGrabar 
         Caption         =   "&Grabar"
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuOpCancelar 
         Caption         =   "&Cancelar"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuSaOut 
         Caption         =   "&Del Formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmTipoFlete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngTipoFlete As Long, douAgenda As Double

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    ObtengoSeteoForm Me
    miBotones False, False, False
    CargoDiasHabilitados
    Exit Sub
errLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inciar el formulario.", Err.Description
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    vsAgenda.Left = 40
    vsAgenda.Width = Me.ScaleWidth - 80
    vsAgenda.Height = Me.ScaleHeight - (vsAgenda.Top + staMensaje.Height + 30)
End Sub

Private Sub CargoDiasHabilitados()
On Error GoTo errCDH
Dim lngCodigo As Long, intCol As Integer, intDia As Integer
Dim rsHora As rdoResultset

    With vsAgenda
        
        Cons = "Select Distinct(HFlNombre), HFlCodigo From HorarioFlete"
        Set rsHora = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
        If rsHora.EOF Then
            .Rows = 0
            .Cols = 0
        Else
            .Rows = 1
            .Cols = 1
            Do While Not rsHora.EOF
                .Cols = .Cols + 1
                .Cell(flexcpText, 0, .Cols - 1) = Trim(rsHora(0))
                lngCodigo = rsHora!HFlCodigo: .Cell(flexcpData, 0, .Cols - 1) = lngCodigo
                .ColAlignment(.Cols - 1) = flexAlignCenterCenter
                rsHora.MoveNext
            Loop
        End If
        rsHora.Close
    
        If .Rows = 0 Then Exit Sub
        
        Cons = "Select * From HorarioFlete Order by HFlDiaSemana"
        Set rsHora = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        intDia = 0
        Do While Not rsHora.EOF
            If intDia <> rsHora!HFlDiaSemana Then
                intDia = rsHora!HFlDiaSemana
                .AddItem DiaSemana(intDia)
                .Cell(flexcpData, .Rows - 1, 0) = intDia
            End If
            intCol = ColumnaHora(rsHora!HFlCodigo)
            lngCodigo = rsHora!HFlIndice: .Cell(flexcpData, .Rows - 1, intCol) = lngCodigo
            .Cell(flexcpChecked, .Rows - 1, intCol) = 2
            .Cell(flexcpBackColor, .Rows - 1, intCol) = vbWindowBackground
            rsHora.MoveNext
        Loop
        rsHora.Close
    End With
    Exit Sub
errCDH:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar la agenda.", Err.Description
End Sub

Private Function ColumnaHora(lngHorario As Long) As Integer
On Error Resume Next
    Dim I As Integer
    ColumnaHora = -1
    For I = 1 To vsAgenda.Cols - 1
        If Val(vsAgenda.Cell(flexcpData, 0, I)) = lngHorario Then ColumnaHora = I: Exit Function
    Next I
End Function

Private Function DiaSemana(ByVal intDia As Integer) As String
    
    Select Case intDia
        Case 1: DiaSemana = "Domingo"
        Case 2: DiaSemana = "Lunes"
        Case 3: DiaSemana = "Martes"
        Case 4: DiaSemana = "Miércoles"
        Case 5: DiaSemana = "Jueves"
        Case 6: DiaSemana = "Viernes"
        Case 7: DiaSemana = "Sábado"
    End Select
    
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    GuardoSeteoForm Me
    CierroConexion
End Sub

Private Sub Label1_Click()
On Error Resume Next
    With tDescripcion
        If .Enabled Then
            .SelStart = 0: .SelLength = Len(.Text): .SetFocus
        End If
    End With
End Sub

Private Sub MnuOpCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuOpGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuOpModificar_Click()
    AccionModificar
End Sub

Private Sub MnuSaOut_Click()
    Unload Me
End Sub

Private Sub tDescripcion_Change()
On Error Resume Next
    LimpioDiasHabilitados
    lngTipoFlete = 0
    miBotones False, False, False
End Sub

Private Sub tDescripcion_GotFocus()
On Error Resume Next
    With tDescripcion
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tDescripcion_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn And Trim(tDescripcion.Text) <> "" Then
        CargoTipoFlete
        If lngTipoFlete > 0 Then
            miBotones True, False, False
        Else
            miBotones False, False, False
        End If
        ArmoAgendaFlete
    End If
End Sub

Private Sub miBotones(bolModificar As Boolean, bolGrabar As Boolean, bolCancelar As Boolean)
    With Toolbar1
        .Buttons("modificar").Enabled = bolModificar
        .Buttons("grabar").Enabled = bolGrabar
        .Buttons("cancelar").Enabled = bolCancelar
    End With
    MnuOpModificar.Enabled = bolModificar
    MnuOpGrabar.Enabled = bolGrabar
    MnuOpCancelar.Enabled = bolCancelar
    
    tDescripcion.Enabled = Not bolGrabar
    vsAgenda.Enabled = bolGrabar
    
End Sub

Private Sub CargoTipoFlete()
On Error GoTo errCTF
    Screen.MousePointer = 11
    
    lngTipoFlete = 0: douAgenda = 0
    
    Cons = "Select Count(*) From TipoFlete Where TFlDescripcion like '" & Replace(tDescripcion.Text, " ", "%") & "%'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux(0) = 0 Then
        RsAux.Close
        MsgBox "No hay tipos de fletes para los datos ingresados.", vbExclamation, "ATENCIÓN"
    ElseIf RsAux(0) > 1 Then
        RsAux.Close
        Cons = "Select tflCodigo, TFlCodigo as Codigo, TFlDescripcion as Descripción From TipoFlete Where TFlDescripcion like '" & Replace(tDescripcion.Text, " ", "%") & "%'"
        Dim objLista As New clsListadeAyuda
        If objLista.ActivarAyuda(cBase, Cons, 4500, 1, "Tipos de Flete") > 0 Then
            lngTipoFlete = objLista.RetornoDatoSeleccionado(0)
        End If
        Set objLista = Nothing
        If lngTipoFlete > 0 Then
            Cons = "Select * From TipoFlete Where TFlCodigo = " & lngTipoFlete
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            tDescripcion.Text = Trim(RsAux!TFlDescripcion)
            lngTipoFlete = RsAux!TFlCodigo
            If Not IsNull(RsAux!TFlAgenda) Then douAgenda = RsAux!TFlAgenda
            RsAux.Close
        End If
    Else
        RsAux.Close
        Cons = "Select * From TipoFlete Where TFlDescripcion like '" & Replace(tDescripcion.Text, " ", "%") & "%'"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        tDescripcion.Text = Trim(RsAux!TFlDescripcion)
        lngTipoFlete = RsAux!TFlCodigo
        If Not IsNull(RsAux!TFlAgenda) Then douAgenda = RsAux!TFlAgenda
        RsAux.Close
    End If
    Screen.MousePointer = 0
    Exit Sub
errCTF:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los tipos de fletes.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    Select Case Button.Key
        Case "modificar": AccionModificar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
    End Select
End Sub

Private Sub LimpioDiasHabilitados()
On Error Resume Next
Dim Fila As Integer, Col As Integer
    For Fila = 1 To vsAgenda.Rows - 1
        For Col = 1 To vsAgenda.Cols - 1
            If Val(vsAgenda.Cell(flexcpData, Fila, Col)) > 0 Then vsAgenda.Cell(flexcpChecked, Fila, Col) = flexUnchecked
        Next Col
    Next Fila
End Sub

Private Sub AccionModificar()
    miBotones False, True, True
    vsAgenda.SetFocus
End Sub

Private Sub AccionCancelar()
On Error Resume Next
    miBotones True, False, False
    tDescripcion.SetFocus
    ArmoAgendaFlete
End Sub

Private Sub ArmoAgendaFlete()
Dim strMat As String, strAux As String
On Error GoTo errSalir

    Screen.MousePointer = 11
    LimpioDiasHabilitados
    
    strMat = superp_MatrizSuperposicion(douAgenda)
    If strMat = "" Then GoTo errSalir
    
    Do While strMat <> ""
        If InStr(1, strMat, ",") > 0 Then
            MarcoEnGrilla CInt(Mid(strMat, 1, InStr(1, strMat, ",") - 1))
            strMat = Mid(strMat, InStr(1, strMat, ",") + 1, Len(strMat))
        Else
            MarcoEnGrilla CInt(strMat)
            strMat = ""
        End If
    Loop
    
errSalir:
    Screen.MousePointer = 0
End Sub

Private Sub MarcoEnGrilla(intIndice As Integer)
On Error Resume Next
Dim Fila As Integer, Col As Integer
    For Fila = 1 To vsAgenda.Rows - 1
        For Col = 1 To vsAgenda.Cols - 1
            If Val(vsAgenda.Cell(flexcpData, Fila, Col)) = intIndice Then vsAgenda.Cell(flexcpChecked, Fila, Col) = flexChecked
        Next Col
    Next Fila
End Sub

Private Sub AccionGrabar()

    If MsgBox("¿Confirma almacenar la agenda ingresada?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
        On Error GoTo errGrabar
        Dim douAux As Double
        Screen.MousePointer = 11
        douAux = CalculoValorSuperposicion
        Cons = "Select * from TipoFlete Where TFlCodigo = " & lngTipoFlete
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.Edit
        RsAux!TFlAgenda = douAux
        RsAux!TFlAgendaHabilitada = douAux
        RsAux!TFlFechaAgeHab = Format(Now, "mm/dd/yyyy hh:mm:ss")
        RsAux.Update
        RsAux.Close
        douAgenda = douAux
        AccionCancelar
        Screen.MousePointer = 0
    End If
    Exit Sub
errGrabar:
    clsGeneral.OcurrioError "Ocurrió un error al intentar almacenar la información.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function CalculoValorSuperposicion() As Double
On Error Resume Next
Dim Fila As Integer, Col As Integer
Dim douAux As Double
    douAux = 0
    For Fila = 1 To vsAgenda.Rows - 1
        For Col = 1 To vsAgenda.Cols - 1
            If vsAgenda.Cell(flexcpChecked, Fila, Col) = flexChecked Then
                douAux = douAux + superp_ValSuperposicion(Val(vsAgenda.Cell(flexcpData, Fila, Col)))
            End If
        Next Col
    Next Fila
    CalculoValorSuperposicion = douAux
End Function

Private Sub vsAgenda_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
    If Val(vsAgenda.Cell(flexcpData, Row, Col)) = 0 Then Cancel = True
End Sub

