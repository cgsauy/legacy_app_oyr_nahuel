VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form frmABM 
   Caption         =   "Agenda Horario de Personal"
   ClientHeight    =   5925
   ClientLeft      =   2970
   ClientTop       =   2310
   ClientWidth     =   8340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmABM.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   8340
   Begin VB.TextBox tbFechaHoy 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   16
      Top             =   480
      Width           =   1455
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsGrid 
      Height          =   1335
      Left            =   120
      TabIndex        =   14
      Top             =   2360
      Width           =   8055
      _ExtentX        =   14208
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
      BackColorBkg    =   -2147483636
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
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
   Begin VB.TextBox tbMemo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      MaxLength       =   200
      TabIndex        =   11
      Top             =   1560
      Width           =   7215
   End
   Begin VB.TextBox tbHora 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      MaxLength       =   5
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ComboBox cbAgenda 
      Height          =   315
      Left            =   960
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1200
      Width           =   2895
   End
   Begin VB.ComboBox cbTipoAgenda 
      Height          =   315
      Left            =   4920
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox tbUsuario 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox tbFecha 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      MaxLength       =   10
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   500
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   5670
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":0DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":10E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":13FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Cambiar fecha hoy:"
      Height          =   255
      Left            =   4920
      TabIndex        =   17
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lbTitulo 
      BackColor       =   &H8000000C&
      Caption         =   "Última Agenda"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   8055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "&Memo:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Hora:"
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "A&genda:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Para:"
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Rige:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuModificar 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuLineEnd 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVolver 
         Caption         =   "&Salir"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmABM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DiAviso As Boolean

Private Function pf_InsertoEnGrilla()
    
    If Not fnc_ValidateSave Then Exit Function
        
    If Val(Toolbar1.Tag) = 0 And Not DiAviso Then
        
        If pf_AgendaUtilizada(tbFecha.Text & " " & tbHora.Text & ":00") Then
            DiAviso = True
            If MsgBox("El usuario ya tiene horarios ingresados para el RIGE que ud ingresó, al grabar afectará posibles consultas." & Chr(13) & Chr(13) & "¿Confirma ingresar la agenda de todas formas?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
                Screen.MousePointer = 0
                Exit Function
            End If
        End If

    End If
    
    If Val(Toolbar1.Tag) = 0 Then
        'Recorro la grilla para buscar si ya tengo insertada este renglón, hago también la validación que si es futura ahí la sostituyo.
        
        Dim iQ As Integer
        Dim iBorrar As Integer
        With vsGrid
            For iQ = 1 To .Rows - 1
                
                If .Cell(flexcpData, iQ, 0) = cbAgenda.ItemData(cbAgenda.ListIndex) And _
                    .Cell(flexcpData, iQ, 1) = cbTipoAgenda.ItemData(cbTipoAgenda.ListIndex) And _
                    .Cell(flexcpData, iQ, 2) = tbFecha.Text & " " & tbHora.Text & ":00" Then
                    
                    .Cell(flexcpText, iQ, 4) = tbMemo.Text
                    pf_InsertoEnGrilla = True
                    Exit Function
                
                ElseIf .Cell(flexcpData, iQ, 0) = cbAgenda.ItemData(cbAgenda.ListIndex) And _
                    .Cell(flexcpData, iQ, 1) = cbTipoAgenda.ItemData(cbTipoAgenda.ListIndex) And _
                    Val(.Cell(flexcpData, iQ, 3)) = 1 Then
                
                    iBorrar = iQ
                    
                End If
            Next
        End With
    End If
    
    With vsGrid
        
        .AddItem Format(tbFecha.Text, "dd/mm/yyyy") & " " & tbHora.Text
        .Cell(flexcpText, .Rows - 1, 1) = cbAgenda.Text
        .Cell(flexcpText, .Rows - 1, 2) = cbTipoAgenda.Text
        .Cell(flexcpText, .Rows - 1, 3) = tbHora.Text
        .Cell(flexcpText, .Rows - 1, 4) = tbMemo.Text
        
        .Cell(flexcpData, .Rows - 1, 0) = cbAgenda.ItemData(cbAgenda.ListIndex)
        .Cell(flexcpData, .Rows - 1, 1) = cbTipoAgenda.ItemData(cbTipoAgenda.ListIndex)
        .Cell(flexcpData, .Rows - 1, 2) = tbFecha.Text & " " & tbHora.Text & ":00"
        
        If CDate(tbFecha.Text & " " & tbHora.Text & ":00") > Now Then
            .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
            '.Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = lColor
            .Cell(flexcpData, .Rows - 1, 4) = 1 'Me digo que es futura
        Else
            .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &H8000&
        End If
        
    End With
    
    If iBorrar > 0 Then vsGrid.RemoveItem iBorrar
        
    pf_InsertoEnGrilla = True
    
End Function

Private Function pf_FindIDEnCombo(ByVal iID As Long, ByVal oCombo As ComboBox) As Integer
Dim iQ As Integer
    pf_FindIDEnCombo = -1
    For iQ = 0 To oCombo.ListCount - 1
        If oCombo.ItemData(iQ) = iID Then
            pf_FindIDEnCombo = iQ
            oCombo.ListIndex = iQ
            Exit Function
        End If
    Next
End Function

Private Function pf_GetIDSAgendaSalida() As String
Dim iQ As Integer
    For iQ = 0 To cbAgenda.ListCount - 1
        pf_GetIDSAgendaSalida = pf_GetIDSAgendaSalida & IIf(pf_GetIDSAgendaSalida = "", "", ", ") & cbAgenda.ItemData(iQ)
    Next
End Function

Private Sub ps_CargoAgenda(ByVal fechaHoy As String)
On Error GoTo errCA
    Dim bFutura As Boolean
    vsGrid.Rows = 1
    Dim lColor As Long: lColor = &HF0F0F0
    Dim dFecha As Date
    Dim sQy As String: sQy = "exec prg_AgendaSalidaHorarioPersonal " & Val(tbUsuario.Tag) & IIf(IsDate(fechaHoy), "," & Format(fechaHoy, "'yyyymmdd'"), ",null")
    Dim rsA As rdoResultset
    'SELECT HPeQue, HPeComentario, HPeTipoAgenda, QHPNombre, TAHNombre, Max(HPeFechaHora)as HPeFechaHora
    Set rsA = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    Do While Not rsA.EOF
        With vsGrid
            
'            If rsA("FechaHora") > Now And Not bFutura Then
'                bFutura = True
'                .AddItem "FUTURAS"
'                .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &H666666
'                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbWhite
'            End If
            
'            If bFutura Then 'And Format(dFecha, "dd/mm/yy") <> Format(rsA("FechaHora"), "dd/mm/yy") Then
'                If lColor = vbWhite Then lColor = &HF0F0F0 Else lColor = vbWhite
'            End If
            
        
            .AddItem Format(rsA("FechaHora"), "dd/mm/yy hh:nn")
            .Cell(flexcpText, .Rows - 1, 1) = Trim(rsA("QueNombre"))
            .Cell(flexcpText, .Rows - 1, 2) = Trim(rsA("TipoAgeNombre"))
            .Cell(flexcpText, .Rows - 1, 3) = Format(rsA("FechaHora"), "hh:nn")
            If Not IsNull(rsA("Comentario")) Then .Cell(flexcpText, .Rows - 1, 4) = Trim(rsA("Comentario"))
            
            .Cell(flexcpData, .Rows - 1, 0) = CStr(rsA("Que"))
            .Cell(flexcpData, .Rows - 1, 1) = CStr(rsA("TipoAgenda"))
            .Cell(flexcpData, .Rows - 1, 2) = CStr(rsA("FechaHora"))
            .Cell(flexcpData, .Rows - 1, 3) = 1 'ME INDICO QUE ESTA INSERTADA POR CARGA.
            
            If rsA("FechaHora") > Now Then
                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue '&H666666
                '.Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = lColor
            End If
            
        End With
        rsA.MoveNext
    Loop
    rsA.Close
    If vsGrid.Rows > 1 Then
        vsGrid.Select 1, 0
        Botones True, True, False
    Else
        Botones True, False, False
    End If
    Exit Sub
errCA:
    clsGeneral.OcurrioError "Error al cargar la agenda del usuario.", Err.Description
End Sub

Private Sub Botones(Nu As Boolean, MoEl As Boolean, Gr As Boolean)

    'Habilito y Desabilito Botones.
    Toolbar1.Buttons("nuevo").Enabled = Nu
    MnuNuevo.Enabled = Nu
    
    Toolbar1.Buttons("modificar").Enabled = MoEl
    MnuModificar.Enabled = MoEl
    
    Toolbar1.Buttons("eliminar").Enabled = MoEl
    MnuEliminar.Enabled = MoEl
    
    Toolbar1.Buttons("grabar").Enabled = Gr
    MnuGrabar.Enabled = Gr
    
    Toolbar1.Buttons("cancelar").Enabled = Gr
    MnuCancelar.Enabled = Gr

End Sub

Private Sub ps_SelectAll(ByVal oCtrl As Control)
On Error Resume Next
    With oCtrl
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub ps_SetFoco(ByVal oCtrl As Control)
On Error Resume Next
    ps_SelectAll oCtrl
    oCtrl.SetFocus
End Sub

Private Sub loc_FillCtrl()
On Error GoTo errFC
    Dim ofncs As New clsFunciones
    ofncs.CargoCombo "SELECT TAHID, TAHNombre FROM TipoAgendaHorarioPersonal ORDER BY TAHPrioridad", cbTipoAgenda
    ofncs.CargoCombo "SELECT * FROM dbo.HorarioSalidasPersonal()", cbAgenda
    Set ofncs = Nothing
    Exit Sub
errFC:
    clsGeneral.OcurrioError "Error al cargar los combos del formulario.", Err.Description, "Cargar combos"
End Sub

Private Sub cbAgenda_GotFocus()
    ps_SelectAll cbAgenda
End Sub

Private Sub cbAgenda_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn And cbAgenda.ListIndex > -1 Then ps_SetFoco tbHora
End Sub

Private Sub cbTipoAgenda_GotFocus()
    ps_SelectAll cbTipoAgenda
End Sub

Private Sub cbTipoAgenda_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn And cbTipoAgenda.ListIndex > -1 Then ps_SetFoco cbAgenda
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
'obtengo de la registry la última posición del formulario.
    Dim ofncs As New clsFunciones
    ofncs.GetPositionForm Me
    Set ofncs = Nothing
    With vsGrid
        .Rows = 1
        .FixedCols = 0: .Cols = 1: .ExtendLastCol = True
        .BackColor = vbWhite: .SheetBorder = vbWhite
        .BackColorBkg = vbWhite
        .FormatString = "Rige|Agenda|Tipo|^Hora|Comentario"
        .ColWidth(0) = 1400
        .ColWidth(1) = 2200
        .ColWidth(2) = 1600
        .ColWidth(3) = 600
        .HighLight = flexHighlightAlways
        .SelectionMode = flexSelectionByRow
    End With
'Inicializo los ctrls
    loc_SetCtrl False
    loc_FillCtrl
    Botones False, False, False
    Screen.MousePointer = 0
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Error al ingresar al formulario."
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
On Error Resume Next
    vsGrid.Move 120, vsGrid.Top, ScaleWidth - 240, ScaleHeight - (vsGrid.Top + Status.Height + 60)
    lbTitulo.Width = vsGrid.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'Guardamos la posición del formulario.
    Dim ofncs As New clsFunciones
    ofncs.SetPositionForm (Me)
    Set ofncs = Nothing
'Cerramos la conexión.
    cBase.Close
'eliminamos la referencia de orcgsa.
    Set clsGeneral = Nothing
    End
    Exit Sub
End Sub

Private Sub Label1_Click()
    ps_SetFoco tbFecha
End Sub

Private Sub Label2_Click()
    ps_SetFoco tbUsuario
End Sub

Private Sub Label3_Click()
    ps_SetFoco cbTipoAgenda
End Sub

Private Sub Label4_Click()
    ps_SetFoco cbAgenda
End Sub

Private Sub Label5_Click()
    ps_SetFoco tbHora
End Sub

Private Sub Label6_Click()
    ps_SetFoco tbMemo
End Sub

Private Sub MnuCancelar_Click()
    loc_CancelarEdicion
End Sub

Private Sub MnuEliminar_Click()
    loc_Eliminar
End Sub

Private Sub MnuGrabar_Click()
    loc_Grabar
End Sub

Private Sub MnuModificar_Click()
    loc_Edicion
End Sub

Private Sub MnuNuevo_Click()
    loc_Nuevo
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Sub loc_Nuevo()
On Error GoTo ErrAN
    Screen.MousePointer = 11
    
'Prendo Señal que es uno nuevo.
    Toolbar1.Tag = 0
    
    'BORRO TODO PARA QUE INGRESE todas juntas.
    'vsGrid.Rows = 1
    
'Habilito y Desabilito Botones.
    Call Botones(False, False, True)

    loc_CleanCtrl
    loc_SetCtrl True
    
    vsGrid.Enabled = True
    
    tbFecha.SetFocus
    Screen.MousePointer = 0
    Exit Sub
    
ErrAN:
    clsGeneral.OcurrioError "Ocurrio un error inesperado.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub loc_Edicion()
    
    If Val(vsGrid.Cell(flexcpData, vsGrid.Row, 0)) = 0 Then
        Exit Sub
    End If
    
'Prendo señal que es modificación.
    Toolbar1.Tag = vsGrid.Row
'Habilito y Desabilito Botones y controles.
    
    Call Botones(False, False, True)
    
    With vsGrid
        tbFecha.Tag = .Cell(flexcpData, .Row, 2)
        cbAgenda.Tag = .Cell(flexcpData, .Row, 0)
        cbTipoAgenda.Tag = .Cell(flexcpData, .Row, 1)
        tbMemo.Text = .Cell(flexcpText, .Row, 3)
    End With
    
    If Not pf_AgendaUtilizada(vsGrid.Cell(flexcpData, vsGrid.Row, 2)) Then
        loc_SetCtrl True
        With vsGrid
            tbFecha.Text = Format(.Cell(flexcpData, .Row, 2), "dd/mm/yyyy")
            
            pf_FindIDEnCombo .Cell(flexcpData, .Row, 0), cbAgenda
            pf_FindIDEnCombo .Cell(flexcpData, .Row, 1), cbTipoAgenda
            
            tbHora.Text = Format(.Cell(flexcpData, .Row, 2), "hh:nn")
            tbMemo.Text = .Cell(flexcpText, .Row, 3)
            
        End With
        tbFecha.SetFocus
    Else
        vsGrid.Enabled = False
        With tbMemo
            .Enabled = True
            .BackColor = vbWindowBackground
            .SetFocus
        End With
    End If
    Screen.MousePointer = 0

End Sub

Private Function pf_AgendaUtilizada(ByVal dRigeDesde As Date) As Boolean
On Error GoTo errAU
Dim sQy As String
Dim rsAux As rdoResultset
    sQy = "SELECT TOP 1 HPeQue FROM HorarioPersonal WHERE HPeUsuario = " & Val(tbUsuario.Tag) & _
        " AND HPeQue NOT IN(" & pf_GetIDSAgendaSalida & ") AND HPeFechaHora > '" & Format(dRigeDesde, "yyyy/mm/dd hh:nn:ss") & "'" & _
        " AND HPeTipoAgenda Is Not Null"
    Set rsAux = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    pf_AgendaUtilizada = Not rsAux.EOF
    rsAux.Close
    Exit Function
errAU:
    clsGeneral.OcurrioError "Error al validar si la agenda fue utilizada.", Err.Description
End Function

Private Sub ps_SaveAllGrid()
On Error GoTo ErrSave

    If MsgBox("Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
        Screen.MousePointer = 11
        
        On Error GoTo errBT
        cBase.BeginTrans
        On Error GoTo errRB
        
        Dim iQ As Integer
        Dim sQy As String
        Dim rsAux As rdoResultset
        
        For iQ = 1 To vsGrid.Rows - 1
            
            If Val(vsGrid.Cell(flexcpData, iQ, 3)) = 0 Then
            
                sQy = "SELECT * FROM HorarioPersonal " & _
                    "WHERE HPeUsuario = " & Val(tbUsuario.Tag) & " AND HPeQue = 0"
                
                Set rsAux = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
                
                rsAux.AddNew
                rsAux("HPeUsuario") = Val(tbUsuario.Tag)
                rsAux("HPeQue") = vsGrid.Cell(flexcpData, iQ, 0)
                rsAux("HPeFechaHora") = Format(vsGrid.Cell(flexcpData, iQ, 2), "yyyy/mm/dd hh:nn:ss")
                If vsGrid.Cell(flexcpText, iQ, 3) = "" Then rsAux("HPeComentario") = Null Else rsAux("HPeComentario") = vsGrid.Cell(flexcpText, iQ, 3)
                rsAux("HPeTipoAgenda") = vsGrid.Cell(flexcpData, iQ, 1)
                rsAux.Update
                rsAux.Close
            End If
        Next
        cBase.CommitTrans
        
    'Invocamos a cancelar p/volver a estado de no edición
        loc_CancelarEdicion
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
    

ErrSave:
    clsGeneral.OcurrioError "No se pudo almacenar la información, reintente.", Err.Description
    Screen.MousePointer = vbDefault
    Exit Sub
errBT:
    clsGeneral.OcurrioError "No se logró iniciar la transacción.", Err.Description
    Screen.MousePointer = vbDefault
    Exit Sub
errRB:
    Resume ROLBACK
    Exit Sub
ROLBACK:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "No se pudo almacenar la información, reintente.", Err.Description
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub loc_Grabar()

    If Val(Toolbar1.Tag) = 0 Then
        ps_SaveAllGrid
        Exit Sub
    End If

'Hacemos los controles de datos ingresados y de validación antes de grabar
    If Not fnc_ValidateSave Then Exit Sub
        
    If MsgBox("Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
        Screen.MousePointer = 11
        
        On Error GoTo ErrSave
        Dim sQy As String
        Dim rsAux As rdoResultset
        
        sQy = "SELECT * FROM HorarioPersonal " & _
            "WHERE HPeUsuario = " & Val(tbUsuario.Tag) & " AND HPeQue = " & Val(cbAgenda.Tag) & " AND HPeTipoAgenda = " & Val(cbTipoAgenda.Tag)
            
        If tbFecha.Tag <> "" Then sQy = sQy & " AND HPeFechaHora = '" & Format(tbFecha.Tag, "yyyy/mm/dd hh:nn:ss") & "'"
        
        Set rsAux = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
        'Si tengo la señal de nuevo
        If rsAux.EOF Then
            rsAux.AddNew
        Else
            rsAux.Edit
        End If
        rsAux("HPeUsuario") = Val(tbUsuario.Tag)
        If cbAgenda.Enabled Then rsAux("HPeQue") = cbAgenda.ItemData(cbAgenda.ListIndex)
        If tbFecha.Enabled Then rsAux("HPeFechaHora") = Format(tbFecha.Text, "yyyy/mm/dd") & " " & tbHora.Text & ":00"
        If tbMemo.Text = "" Then rsAux("HPeComentario") = Null Else rsAux("HPeComentario") = Trim(tbMemo.Text)
        If cbTipoAgenda.Enabled Then rsAux("HPeTipoAgenda") = cbTipoAgenda.ItemData(cbTipoAgenda.ListIndex)
        rsAux.Update
        rsAux.Close
        
    'Invocamos a cancelar p/volver a estado de no edición
        loc_CancelarEdicion
    End If
    Exit Sub
    

ErrSave:
    clsGeneral.OcurrioError "No se pudo almacenar la información, reintente.", Err.Description
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub loc_Eliminar()
On Error GoTo errDel
    'Verificar si hay datos a validar.
    Dim sMsg As String
    Dim iOpt As Byte
    
    If pf_AgendaUtilizada(vsGrid.Cell(flexcpData, vsGrid.Row, 2)) Then
        iOpt = 1
        sMsg = "Confrima dejar sin uso la agenda seleccionada?"
    Else
        iOpt = 2
        sMsg = "Confrima eliminar la agenda seleccionada?"
    End If
    
    
    If MsgBox(sMsg, vbQuestion + vbYesNo, "Eliminar") = vbYes Then
        Screen.MousePointer = 11
        Dim sQy As String
        Dim rsAux As rdoResultset
        sQy = "SELECT * FROM HorarioPersonal" & _
            " WHERE HPeUsuario = " & Val(tbUsuario.Tag) & " AND HPeQue = " & Val(vsGrid.Cell(flexcpData, vsGrid.Row, 0)) & " AND HPeTipoAgenda = " & Val(vsGrid.Cell(flexcpData, vsGrid.Row, 1)) & _
            " AND HPeFechaHora = '" & Format(vsGrid.Cell(flexcpData, vsGrid.Row, 2), "yyyy/mm/dd hh:nn:ss") & "'"
        Set rsAux = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            If iOpt = 2 Then
                rsAux.Delete
            Else
                rsAux.Edit
                'rsAux("HPeFechaHora") = Format(vsGrid.Cell(flexcpData, vsGrid.Row, 2), "yyyy/mm/dd 01:11:00")
                rsAux("HPeEnUso") = 0
                rsAux.Update
            End If
        End If
        rsAux.Close
        'Limpiamos los controles y ponemos el formulario en su nuevo estado.
        ps_CargoAgenda ""
        'Botones True, False, False
        Screen.MousePointer = 0
    End If
    Exit Sub
errDel:
    clsGeneral.OcurrioError "No se pudo eliminar el registro.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub loc_CancelarEdicion()
    Screen.MousePointer = 11
    loc_SetCtrl False
    loc_CleanCtrl
'Si es edición cargamos los valores si no limpiamos la ficha.
    ps_CargoAgenda ""
    'Elimino señal de edición.
    Toolbar1.Tag = 0
    Screen.MousePointer = 0
End Sub


Private Sub tbFecha_Change()
    If Val(Toolbar1.Tag) = 0 Then DiAviso = False
End Sub

Private Sub tbFecha_GotFocus()
    ps_SelectAll tbFecha
End Sub

Private Sub tbFecha_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If IsDate(tbFecha.Text) Then
            tbFecha.Text = Format(tbFecha.Text, "dd/mm/yyyy")
            cbTipoAgenda.SetFocus
        End If
    End If
End Sub


Private Sub tbFechaHoy_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And Val(tbUsuario.Tag) > 0 Then
        loc_CleanCtrl
        vsGrid.Rows = 1
        If IsDate(tbFechaHoy.Text) Then
            ps_CargoAgenda tbFechaHoy.Text
            If vsGrid.Rows > 1 Then Botones False, False, False
        End If
    End If
    
End Sub

Private Sub tbHora_GotFocus()
    ps_SelectAll tbHora
End Sub

Private Sub tbHora_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn And tbHora.Text <> "" Then
        tbHora.Text = Replace(tbHora.Text, ".", ":")
        Select Case Len(tbHora.Text)
            Case 1, 2
                If IsNumeric(tbHora.Text) Then
                    tbHora.Text = Format(tbHora.Text, "00") & ":00"
                Else
                    MsgBox "No se interpreto la hora", vbExclamation, "Atención"
                    Exit Sub
                End If
            Case Is > 2
                If InStr(1, tbHora.Text, ":") > 0 Then
                    If InStr(1, tbHora.Text, ":") = 1 Then
                        tbHora.Text = "00:" & Format(Val(Mid(tbHora.Text, 2)), "00")
                    Else
                        If InStr(1, tbHora.Text, ":") = 2 And Len(tbHora.Text) = 3 Then
                            tbHora.Text = "0" & tbHora.Text & "0"
                        ElseIf InStr(1, tbHora.Text, ":") = 3 And Len(tbHora.Text) = 4 Then
                            tbHora.Text = Format(Val(Mid(tbHora.Text, 1, InStr(1, tbHora.Text, ":") - 1)), "00") & ":" & Mid(tbHora.Text, InStr(1, tbHora.Text, ":") + 1) & "0"
                        Else
                            tbHora.Text = Format(Val(Mid(tbHora.Text, 1, InStr(1, tbHora.Text, ":") - 1)), "00") & ":" & Format(Val(Mid(tbHora.Text, InStr(1, tbHora.Text, ":") + 1)), "00")
                        End If
                    End If
                Else
                    tbHora.Text = Format(tbHora.Text, "0000")
                    tbHora.Text = Mid(tbHora.Text, 1, 2) & ":" & Mid(tbHora.Text, 3)
                End If
        End Select
        ps_SetFoco tbMemo
    End If
End Sub

Private Sub tbMemo_GotFocus()
    ps_SelectAll tbMemo
End Sub

Private Sub tbMemo_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If Val(Toolbar1.Tag) = 0 Then
            'Inserto en grilla.
            If pf_InsertoEnGrilla Then
                'Dejo solo la fecha
                cbTipoAgenda.Text = ""
                cbAgenda.Text = ""
                tbMemo.Text = ""
                tbHora.Text = ""
                cbTipoAgenda.SetFocus
            End If
        Else
            loc_Grabar
        End If
    End If
End Sub

Private Sub tbUsuario_Change()
    If Val(tbUsuario.Tag) > 0 Then
        vsGrid.Rows = 1
        tbUsuario.Tag = 0
        Botones False, False, False
    End If
End Sub

Private Sub tbUsuario_GotFocus()
    ps_SelectAll tbUsuario
End Sub

Private Sub tbUsuario_KeyPress(KeyAscii As Integer)
On Error GoTo errFU
    If KeyAscii = vbKeyReturn Then
        If Val(tbUsuario.Tag) = 0 And tbUsuario.Text <> "" Then
            Dim sQy As String
            Dim rsAux As rdoResultset
            sQy = "Select UsuCodigo, UsuIdentificacion Identificación, UsuNombre1 + IsNull(UsuApellido1, '') Nombre  From Usuario Where UsuDigito = " & IIf(Val(tbUsuario.Text) > 0, Val(tbUsuario.Text), -8888) & " OR UsuIdentificacion Like '" & Replace(tbUsuario.Text, " ", "%") & "%' And UsuHabilitado = 1"
            Set rsAux = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
            If Not rsAux.EOF Then
                tbUsuario.Text = Trim(rsAux!Identificación)
                tbUsuario.Tag = rsAux(0)
                rsAux.MoveNext
                If Not rsAux.EOF Then
                    Dim objAyuda As New clsListadeAyuda
                    If objAyuda.ActivarAyuda(cBase, sQy, 4000, 1, "Usuario") > 0 Then
                        tbUsuario.Text = objAyuda.RetornoDatoSeleccionado(1)
                        tbUsuario.Tag = objAyuda.RetornoDatoSeleccionado(0)
                    End If
                    Set objAyuda = Nothing
                End If
                
            Else
                MsgBox "No hay un usuario con esos datos.", vbInformation, "Buscar usuario"
                ps_SelectAll tbUsuario
            End If
            rsAux.Close
            Botones Val(tbUsuario.Tag) > 0, False, False
            If Val(tbUsuario.Tag) > 0 Then ps_CargoAgenda ""
        End If
    End If
    Exit Sub
errFU:
    clsGeneral.OcurrioError "Error al buscar el usuario.", Err.Description, Me.Caption
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": loc_Nuevo
        Case "modificar": loc_Edicion
        Case "eliminar": loc_Eliminar
        Case "grabar": loc_Grabar
        Case "cancelar": loc_CancelarEdicion
        Case "salir": Unload Me
    End Select

End Sub

Private Function fnc_ValidateSave() As Boolean
'Arrancamos la función en false (valor x defecto) para salir al primer error que encontremos.
    fnc_ValidateSave = False
    
    If Not IsDate(tbFecha.Text) And tbFecha.Enabled Then
        MsgBox "La fecha ingresada no es válida.", vbExclamation, "Validación"
        ps_SetFoco tbFecha
        Exit Function
    End If
    
    If cbTipoAgenda.ListIndex = -1 And cbTipoAgenda.Enabled Then
        MsgBox "Seleccione para que días ocurre la agenda.", vbExclamation, "Validación"
        ps_SetFoco cbTipoAgenda
        Exit Function
    End If
    
    If cbAgenda.ListIndex = -1 And cbAgenda.Enabled Then
        MsgBox "De indicar cual es la agenda que esta ingresando.", vbExclamation, "Validación"
        ps_SetFoco cbAgenda
        Exit Function
    End If
    
    If tbHora.Enabled Then
        If InStr(1, tbHora.Text, ":") = 0 Then
            MsgBox "La hora no es correcta.", vbExclamation, "Validación"
            ps_SetFoco tbHora
            Exit Function
        ElseIf Not IsNumeric(Replace(tbHora.Text, ":", "")) Then
            MsgBox "La hora no es correcta.", vbExclamation, "Validación"
            ps_SetFoco tbHora
            Exit Function
        End If
        
        If tbHora.Text = "01:11" Then
            MsgBox "Al grabar la hora 01:11 ud está dejando inactiva la agenda.", vbExclamation, "Atención"
        End If
    End If
    
    'Recorro la grilla para validar que no tenga la misma ya ingresada.
    If tbFecha.Enabled Then
        Dim iQ As Integer
        For iQ = 1 To vsGrid.Rows - 1
            If Val(Toolbar1.Tag) <> iQ And _
                    cbAgenda.ItemData(cbAgenda.ListIndex) = vsGrid.Cell(flexcpData, iQ, 0) And cbTipoAgenda.ItemData(cbTipoAgenda.ListIndex) = vsGrid.Cell(flexcpData, iQ, 1) Then
                'CDate(vsGrid.Cell(flexcpData, vsGrid.Row, 2))
'                If CDate(Format(tbFecha.Text, "dd/mm/yyyy") & " " & tbHora.Text & ":00") < Date And vsGrid.Cell(flexcpForeColor, iQ, 0) = 0 Then
 '                   MsgBox "La fecha es menor a hoy y ya ingresó una agenda con el mismo tipo, duplicación de datos.", vbExclamation, "Atención"
  '                  Exit Function
'                Else
'                    If CDate(Format(tbFecha.Text, "dd/mm/yyyy") & " " & tbHora.Text & ":00") > Date And vsGrid.Cell(flexcpForeColor, iQ, 0) <> 0 Then
'                        MsgBox "Ya ingresó una agenda con el mismo tipo, duplicación de datos.", vbExclamation, "Atención"
'                        Exit Function
'                    End If
   '             End If
                
            End If
        Next
    End If
    fnc_ValidateSave = True
    
End Function

Private Sub loc_SetCtrl(ByVal bEdit As Boolean)
'Rutina para habilitar/deshabilitar los controles
    tbUsuario.Enabled = Not bEdit
    tbUsuario.BackColor = IIf(Not bEdit, vbWindowBackground, vbButtonFace)
    vsGrid.Enabled = Not bEdit
    tbFechaHoy.Enabled = Not bEdit
    tbFechaHoy.BackColor = tbUsuario.BackColor
   
    With tbFecha
        .Enabled = bEdit
        .BackColor = IIf(bEdit, vbWindowBackground, vbButtonFace)
    End With
    
    With cbTipoAgenda
        .Enabled = bEdit
        .BackColor = IIf(bEdit, vbWindowBackground, vbButtonFace)
    End With
    With cbAgenda
        .Enabled = bEdit
        .BackColor = IIf(bEdit, vbWindowBackground, vbButtonFace)
    End With
    
    With tbHora
        .Enabled = bEdit
        .BackColor = IIf(bEdit, vbWindowBackground, vbButtonFace)
    End With
    
    With tbMemo
        .Enabled = bEdit
        .BackColor = IIf(bEdit, vbWindowBackground, vbButtonFace)
    End With
    
End Sub

Private Sub loc_CleanCtrl()
'limpiamos los controles con sus valores x defecto.
    tbFecha.Text = "": tbFecha.Tag = ""
    cbTipoAgenda.Text = ""
    cbAgenda.Text = ""
    tbMemo.Text = ""
    tbHora.Text = ""
End Sub

Private Sub vsGrid_DblClick()
    
    If vsGrid.Rows = 1 Then Exit Sub
    
    If MnuModificar.Enabled And vsGrid.Row > 0 Then
        
        loc_Edicion
    
    ElseIf vsGrid.Row > 0 Then
        
        If (Val(Toolbar1.Tag) = 0 And Val(vsGrid.Cell(flexcpData, vsGrid.Row, 3)) = 0) Then
            
            'Edito
            With vsGrid
            
                tbFecha.Tag = .Cell(flexcpData, .Row, 2)
                tbFecha.Text = Format(.Cell(flexcpData, .Row, 2), "dd/mm/yyyy")
                cbAgenda.Tag = .Cell(flexcpData, .Row, 0)
                cbTipoAgenda.Tag = .Cell(flexcpData, .Row, 1)
                tbMemo.Text = .Cell(flexcpText, .Row, 3)
                pf_FindIDEnCombo .Cell(flexcpData, .Row, 0), cbAgenda
                pf_FindIDEnCombo .Cell(flexcpData, .Row, 1), cbTipoAgenda
                tbMemo.Text = .Cell(flexcpText, .Row, 3)
                tbHora.Text = Format(.Cell(flexcpData, .Row, 2), "hh:nn")
                
                .RemoveItem .Row
                
                cbTipoAgenda.SetFocus
                
            End With
        
        End If
    End If
End Sub

Private Sub vsGrid_RowColChange()
On Error Resume Next
    
    If Not MnuGrabar.Enabled Then
        
        If MnuNuevo.Enabled And vsGrid.Row > 0 Then
            Botones True, Val(vsGrid.Cell(flexcpData, vsGrid.Row, 0)) > 0, False
        ElseIf MnuNuevo.Enabled Then
            Botones Val(tbUsuario.Tag) > 0, False, False
        End If
        
    End If
End Sub

