VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmModificarHorario 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Modificar horario de personal"
   ClientHeight    =   5580
   ClientLeft      =   1485
   ClientTop       =   1770
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmModificarHorario.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   9480
   Begin VB.TextBox txtUsuario 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   1058
      ButtonWidth     =   2540
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cerrar"
            Key             =   "close"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            Key             =   "query"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Transitorias"
            Key             =   "transitoria"
            Object.ToolTipText     =   "Agregar nuevas columnas ttransitorias"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Agregar"
            Key             =   "new"
            Object.ToolTipText     =   "Un día no existente en la grilla"
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   6
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "entrada"
                  Text            =   "Entrada"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "almuerzoS"
                  Text            =   "Salió a almorzar"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "almuerzoR"
                  Text            =   "Retorno almuerzo"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "salida"
                  Text            =   "Salida"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "transitoriaS"
                  Text            =   "Salio transitoriamente"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "transitoriaR"
                  Text            =   "Regreso transitorio"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ausencias"
            Key             =   "ausencias"
            Object.ToolTipText     =   "Ausencias y otras causas"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.TextBox tFecha 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      TabIndex        =   3
      Text            =   "88/88/8888"
      Top             =   720
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8400
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmModificarHorario.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmModificarHorario.frx":11DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmModificarHorario.frx":1AB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmModificarHorario.frx":1F8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmModificarHorario.frx":2577
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8LCtl.VSFlexGrid vsQuery 
      Height          =   3615
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   8415
      _cx             =   14843
      _cy             =   6376
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
      Rows            =   50
      Cols            =   10
      FixedRows       =   0
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
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   2
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lbFecha 
      BackStyle       =   0  'Transparent
      Caption         =   "&Fecha:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "frmModificarHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const cte_MemoUpdate As String = "[Usr=[Digito], HrAnt=[hora], FMod=[now], Def=[defi]]"
Private Const cte_BuscarMemo As String = "[Usr="

Private Enum colGrillaHorario
    chFecha = 0
'    chEntrada
'    chAlmuerzoS
'    chAlmuerzoR
'    chSalida
'    chTransitoriasHoras
'    chTransitoriasTiempo
'    chTotalHoras
    chAccion
    chComentario
End Enum


Private usuDigito As Integer

Private Sub AgregarEvento(ByVal Que As HorariosFijos)
    
    Dim frmQue As New frmAusencias
    If IsDate(tFecha.Text) Then
        frmQue.Dia = CDate(tFecha.Text)
    Else
        frmQue.Dia = Date
    End If
    frmQue.Que = Que
    frmQue.Usuario = Val(txtUsuario.Tag)
    frmQue.Show vbModal
    CargarHorasUsuario
    Exit Sub
   
   
'    Dim sDia As String
'    sDia = InputBox("Ingrese la hora en que se realizó la acción", "Agregar acción", Format(Now, "HH:nn"))
'    sDia = StringAHora(sDia)
'    If IsDate(sDia) Then
'        'Hago validación de si existe o no ese ingreso.
'        sDia = Format(CDate(tFecha.Text & " " & sDia), "dd/mm/yyyy hh:nn:00")
'        Dim sDef As String
'        sDef = InputBox("Ingrese un comentario que explique la acción.", "Ingresar acción")
'
'        'Private Const cte_MemoUpdate As String = "Usr=[Digito], HrAnt=[hora], FMod=[now], Def=[defi]"
'        Dim sMemo As String
'        sMemo = Replace(Replace(cte_MemoUpdate, "[Digito]", usuDigito), "[hora]", "")
'        sMemo = Replace(Replace(sMemo, "[now]", Now), "[defi]", sDef)
'
'        Dim bIng As Boolean
'
'        On Error GoTo errVD
'        Dim rsU As rdoResultset
'        Dim cons As String
'        cons = "SELECT * FROM HorarioPersonal WHERE HPeUsuario = " & Val(txtUsuario.Tag) _
'            & " AND HPeFechaHora between '" & Format(sDia, "yyyy/mm/dd 00:00:00") & "' AND '" & Format(sDia, "yyyy/mm/dd 23:59:59") & "' " _
'            & " AND HPeTipoAgenda Is Null AND HPeQue = " & Que
'        Set rsU = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
'        If rsU.EOF Or Que = TransitoriaS Or Que = TransitoriaR Then
'            rsU.AddNew
'            rsU("HPeUsuario") = Val(txtUsuario.Tag)
'            rsU("HPeQue") = HorariosFijos.Ingreso
'            rsU("HPeFechaHora") = Format(sDia, "yyyy/mm/dd hh:nn:00")
'            rsU("HPeComentario") = sMemo
'            rsU.Update
'            bIng = True
'        Else
'            MsgBox "Ya existe un registro para el día " & sDia & ", debe editar en la grilla.", vbExclamation, "ATENCIÓN"
'        End If
'        rsU.Close
'        If bIng Then CargarHorasUsuario
'    End If
'    Exit Sub
errVD:
    clsGeneral.OcurrioError "Error al ingresar la fecha", Err.Description, "Agregar un día"
End Sub

Private Sub CargarHorasUsuario()
On Error GoTo errCA
    
    InicializoGrillas
    vsQuery.Redraw = False
    
    If Val(txtUsuario.Tag) = 0 Or Not IsDate(tFecha.Text) Then
        MsgBox "Los filtros son obligatorios.", vbInformation, "Faltan filtros"
        Exit Sub
    End If
    tFecha.Tag = tFecha.Text
        
    Dim Dia As Date
    Dia = DateAdd("m", -1, CDate(tFecha.Text))
    Dim Padre As Integer
    
    Dim cons As String
    cons = "SELECT * FROM HorarioPersonal INNER JOIN QueHorarioPersonal ON QHPID = HPeQue " _
        & "WHERE HPeUsuario = " & Val(txtUsuario.Tag) _
        & " AND HPeFechaHora Between '" & Format(CDate(tFecha.Text), "yyyy/mm/dd 00:00:00") & "' AND '" _
        & Format(CDate(tFecha.Text), "yyyy/mm/dd 23:59:59") & "' AND HPeTipoAgenda IS NULL " _
        & "ORDER BY HPeFechaHora"
    Dim rsH As rdoResultset
    Screen.MousePointer = 11
    Set rsH = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsH.EOF
        With vsQuery
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, colGrillaHorario.chFecha) = Format(rsH("HPeFechaHora"), "HH:nn")
            .Cell(flexcpData, .Rows - 1, chFecha) = CStr(rsH("HPeFechaHora"))
            .Cell(flexcpText, .Rows - 1, chAccion) = Trim(rsH("QHPNombre"))
            .Cell(flexcpData, .Rows - 1, chAccion) = CStr(rsH("HPeQue"))
            If Not IsNull(rsH("HPeComentario")) Then .Cell(flexcpText, .Rows - 1, chComentario) = Trim(rsH("HPeComentario"))
            
            If rsH("HPeQUE") >= 30 Then
                .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = &H8000000F
            End If
        End With
        rsH.MoveNext
    Loop
    rsH.Close
    vsQuery.Redraw = True
    Screen.MousePointer = 0
    Exit Sub
errCA:
    Screen.MousePointer = 0
    vsQuery.Redraw = True
    clsGeneral.OcurrioError "Error al cargar las horas del usuario.", Err.Description, "Cargar horas"
End Sub

Private Function GetQueryFind(ByVal sTxt As String) As String
    GetQueryFind = "'" & Replace(Replace(sTxt, "'", "''"), " ", "%") & "%'"
End Function

Private Sub InicializoGrillas()

    On Error Resume Next
        
    With vsQuery
        .FixedRows = 1
        .Rows = .FixedRows
        .Editable = True
        
        .FixedCols = 0
        .Cols = 1
        .GridLinesFixed = flexGridFlatHorz
        .GridLines = flexGridFlatHorz
        
        .FormatString = "<Hora|Acción|Comentario"
        .ColWidth(colGrillaHorario.chFecha) = 800
        .ColWidth(colGrillaHorario.chAccion) = 2100
        .ColWidth(colGrillaHorario.chComentario) = 3100
        
        .ColAlignment(chFecha) = flexAlignCenterCenter
        
        .ExtendLastCol = True
        .WordWrap = False
               
        .BackColorBkg = .BackColor
        '.GridLinesFixed = flexGridInsetHorz
        .GridLinesFixed = flexGridFlat
        
        .BackColorFixed = &H9C6A58  'vbApplicationWorkspace
        .ForeColorFixed = &HFAF0EB  'vbWindowBackground
        .GridColorFixed = &HFAF0EB
        
        .HighLight = flexHighlightAlways ' flexHighlightWithFocus
        .FocusRect = flexFocusNone
        .BackColorSel = &HD3F5F9 '&HC08888       '&HD3F5F9    'vbInfoBackground
        .ForeColorSel = vbBlack '1 'vbHighlight
        .SheetBorder = .BackColor
        
        '.GridColorFixed = vbButtonShadow
        '.ForeColorFixed = vbHighlightText
'        .BorderStyle = flexBorderNone
        .MergeCells = flexMergeSpill
        
'        .SelectionMode = 1
        
        .RowHeight(0) = 320
        .RowHeightMin = 285
        
        .OutlineCol = 0
        
        .AllowUserResizing = flexResizeColumns
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSizeMouse = True
        .AllowSelection = False
    End With
      
End Sub

Private Sub Form_Load()
On Error Resume Next
        
    tFecha.Text = Date
    MenuToolbar
    
    Dim rsU As rdoResultset
    Set rsU = cBase.OpenResultset("SELECT UsuDigito FROM Usuario WHERE UsuCodigo =" & usrCodigo, rdOpenDynamic, rdConcurValues)
    If Not rsU.EOF Then
        usuDigito = rsU(0)
    Else
        MsgBox "No cargué el dígito del usuario logueado", vbExclamation, "ATENCIÓN"
    End If
    rsU.Close

    InicializoGrillas
    Screen.MousePointer = 0
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    Screen.MousePointer = 11
    
    With vsQuery
        .Left = 60
        .Top = txtUsuario.Top + txtUsuario.Height + 90
        .Height = Me.ScaleHeight - (.Top + 30)
        .Width = Me.ScaleWidth - (.Left * 2)
    End With
    Screen.MousePointer = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set clsGeneral = Nothing
    cBase.Close
End Sub

Private Sub Label3_DblClick()
On Error Resume Next
    vsQuery.SaveGrid "grilla.xls", flexFileExcel
End Sub

Private Sub tFecha_Change()
    If tFecha.Tag <> "" Then
        InicializoGrillas
        tFecha.Tag = ""
    End If
End Sub

Private Sub tFecha_GotFocus()
    With tFecha
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
On Error GoTo errKP
    If KeyAscii = vbKeyReturn Then
        If IsDate(tFecha.Text) And tFecha.Tag = "" Then
            tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy")
            tFecha.Tag = tFecha.Text
        End If
        If IsDate(tFecha.Text) Then CargarHorasUsuario
    End If
errKP:
End Sub

Private Sub tFecha_LostFocus()
    If Not IsDate(tFecha.Text) Then tFecha.Text = "" Else tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy")
End Sub

Private Sub tHasta_Change()
    If tFecha.Tag <> "" Then
        InicializoGrillas
        tFecha.Tag = ""
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "close": Unload Me
'        Case "print": loc_ActionPrint True
'        Case "preview": loc_ShowPreview
        Case "query": CargarHorasUsuario
        'Case "transitoria": AgregoColTransitorias
'        Case "new"
'            If Val(txtUsuario.Tag) > 0 Then
'                AgregarUnaFecha
'            End If
        Case "ausencias"
            Dim frmAus As New frmAusencias
            frmAus.Usuario = Val(txtUsuario.Tag)
            frmAus.Show vbModal
            CargarHorasUsuario
    End Select
End Sub

Private Sub MenuToolbar()
    Toolbar1.Buttons("query").Enabled = (Val(txtUsuario.Tag) > 0)
    Toolbar1.Buttons("new").Enabled = (Val(txtUsuario.Tag) > 0)
    Toolbar1.Buttons("ausencias").Enabled = (Val(txtUsuario.Tag) > 0)
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Debug.Print ButtonMenu.Key
    Select Case LCase(ButtonMenu.Key)
        Case "entrada": AgregarEvento Ingreso
        Case "almuerzos": AgregarEvento AlmuerzoS
        Case "almuerzor": AgregarEvento AlmuerzoR
        Case "salida": AgregarEvento Salida
        Case "transitoriar": AgregarEvento TransitoriaR
        Case "transitorias": AgregarEvento TransitoriaS
    End Select
End Sub

Private Sub txtUsuario_Change()
    If Val(txtUsuario.Tag) > 0 Then
        InicializoGrillas
        txtUsuario.Tag = ""
        MenuToolbar
    End If
End Sub

Private Sub txtUsuario_GotFocus()
    With txtUsuario
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
On Error GoTo errTU
    
    If KeyAscii = vbKeyReturn And Trim(txtUsuario.Text) <> "" Then
        If Val(txtUsuario.Tag) = 0 Then
            Screen.MousePointer = 11
            Dim cons As String
            cons = "SELECT UsuCodigo, UsuDigito Digito, RTrim(UsuIdentificacion) Identificación " _
                & "FROM Usuario "
                
            If IsNumeric(txtUsuario.Text) Then
                cons = cons & "WHERE UsuDigito = " & txtUsuario.Text
            Else
                cons = cons & "WHERE UsuIdentificacion Like " & GetQueryFind(txtUsuario.Text) _
                & "ORDER BY UsuIdentificacion"
            End If
            Dim rsU As rdoResultset
            Set rsU = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            If Not rsU.EOF Then
                rsU.MoveNext
                If rsU.EOF Then
                    rsU.MoveFirst
                    txtUsuario.Text = Trim(rsU("Identificación"))
                    txtUsuario.Tag = rsU("UsuCodigo")
                Else
                    Dim oAyuda As New clsListadeAyuda
                    If oAyuda.ActivarAyuda(cBase, cons, 4500, 1, "Búsqueda de usuario") > 0 Then
                        txtUsuario.Text = oAyuda.RetornoDatoSeleccionado(2)
                        txtUsuario.Tag = oAyuda.RetornoDatoSeleccionado(0)
                    End If
                    Set oAyuda = Nothing
                End If
                rsU.Close
            Else
                MsgBox "No se encontró un usuario para el filtro ingresado.", vbInformation, "Búsqueda"
            End If
            Screen.MousePointer = 0
            MenuToolbar
            If Val(txtUsuario.Tag) > 0 Then tFecha.SetFocus
        Else
            tFecha.SetFocus
        End If
    End If
    Exit Sub
errTU:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al buscar el usuario.", Err.Description, "Usuario"
End Sub

Private Sub vsQuery_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo errS
    
    'If (vsQuery.Cell(flexcpText, Row, Col) <> Format(vsQuery.Cell(flexcpData, Row, Col), "HH:nn") Or Trim(vsQuery.Cell(flexcpData, Row, Col)) = "") Then
        
        
        If CStr(vsQuery.Cell(flexcpData, Row, Col)) = "" Then
            MsgBox "Mal seteo de hora, vuelva a cargar la grilla.", vbInformation, "ATENCIÓN"
            Exit Sub
        End If
        
        Dim sDef As String
        sDef = InputBox("Ingrese un comentario que explique la modificación", "Modificar horario")
        
        Dim rsU As rdoResultset
        Dim cons As String
        cons = "SELECT * FROM HorarioPersonal WHERE HPeUsuario = " & Val(txtUsuario.Tag) _
            & " AND HPeFechaHora = '" & Format(vsQuery.Cell(flexcpData, Row, Col), "yyyy/mm/dd HH:nn:ss") & "'" _
            & " AND HPeTipoAgenda Is Null AND HPeQue = " & vsQuery.Cell(flexcpData, Row, colGrillaHorario.chAccion)
            
        Set rsU = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        
        'Private Const cte_MemoUpdate As String = "Usr=[Digito], HrAnt=[hora], FMod=[now], Def=[defi]"
        Dim sMemo As String
        sMemo = Replace(Replace(cte_MemoUpdate, "[Digito]", usuDigito), "[hora]", Format(vsQuery.Cell(flexcpData, Row, Col), "HH.nn"))
        sMemo = Replace(Replace(sMemo, "[now]", Now), "[defi]", sDef)
            
        If Not rsU.EOF Then
            rsU.Edit
            rsU("HPeFechaHora") = Format(tFecha.Text, "yyyy/mm/dd ") & vsQuery.Cell(flexcpText, Row, Col) & ":00"
            Dim sMemoAnt As String
            If InStr(1, rsU("HPeComentario"), cte_BuscarMemo, vbTextCompare) > 0 Then
                sMemoAnt = Mid(rsU("HPeComentario"), InStr(1, rsU("HPeComentario"), cte_BuscarMemo, vbTextCompare))
                sMemoAnt = Mid(sMemoAnt, 1, InStr(2, sMemoAnt, "]"))
                sMemo = sMemo & Replace(rsU("HPeComentario"), sMemoAnt, "")
            Else
                sMemo = sMemo & rsU("HPeComentario")
            End If
            rsU("HPeComentario") = sMemo
            rsU.Update
        Else
            MsgBox "No encontré el registro a editar, verifique que otro usuario no haya modificado la información.", vbExclamation, "ATENCIÓN"
        End If
        rsU.Close
        
        Dim dateBuscar As Date
        dateBuscar = vsQuery.TextMatrix(Row, 0)
        CargarHorasUsuario
    'End If
    Exit Sub
errS:
    clsGeneral.OcurrioError "Error al editar.", Err.Description, "Edición de horas"
End Sub

Private Sub vsQuery_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col <> colGrillaHorario.chFecha Or vsQuery.Cell(flexcpBackColor, Row, Col) = &H8000000F)
End Sub

Private Sub vsQuery_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errdelete
    If KeyCode = vbKeyDelete Then
        If vsQuery.Row >= 1 Then
            If MsgBox("¿Confirma eliminar el registro seleccionado?", vbQuestion + vbYesNo, "Eliminar hora") = vbYes Then
                Dim sQry As String
                sQry = "DELETE logdb.dbo.HorarioPersonal WHERE HPeUsuario = " & Val(txtUsuario.Tag) & " AND HPeQue = " & vsQuery.Cell(flexcpData, vsQuery.Row, chAccion) & _
                     " AND HPeFechaHora = '" & Format(vsQuery.Cell(flexcpData, vsQuery.Row, chFecha), "yyyyMMdd HH:nn:ss") & "'"
                cBase.Execute (sQry)
                CargarHorasUsuario
            End If
        End If
    End If
    Exit Sub
errdelete:
    clsGeneral.OcurrioError "Error al eliminar", Err.Description
End Sub

Private Sub vsQuery_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    vsQuery.EditText = Replace(vsQuery.EditText, ".", ":")
    If Len(vsQuery.EditText) > 5 Then
        Cancel = True
        MsgBox "El formato a ingresar es Hora:Minutos (##;## o ####)", vbExclamation, "Atención"
    Else
        Select Case Len(vsQuery.EditText)
            Case 1, 2
                If IsNumeric(vsQuery.EditText) Then
                    vsQuery.EditText = Format(vsQuery.EditText, "00") & ":00"
                Else
                    MsgBox "No se interpreto la hora", vbExclamation, "Atención"
                    Cancel = True
                    Exit Sub
                End If
            Case Is > 2
                If InStr(1, vsQuery.EditText, ":") > 0 Then
                    If InStr(1, vsQuery.EditText, ":") = 1 Then
                        vsQuery.EditText = "00:" & Format(Val(Mid(vsQuery.EditText, 2)), "00")
                    Else
                        If InStr(1, vsQuery.EditText, ":") = 2 And Len(vsQuery.EditText) = 3 Then
                            vsQuery.EditText = "0" & vsQuery.EditText & "0"
                        ElseIf InStr(1, vsQuery.EditText, ":") = 3 And Len(vsQuery.EditText) = 4 Then
                            vsQuery.EditText = Format(Val(Mid(vsQuery.EditText, 1, InStr(1, vsQuery.EditText, ":") - 1)), "00") & ":" & Mid(vsQuery.EditText, InStr(1, vsQuery.EditText, ":") + 1) & "0"
                        Else
                            vsQuery.EditText = Format(Val(Mid(vsQuery.EditText, 1, InStr(1, vsQuery.EditText, ":") - 1)), "00") & ":" & Format(Val(Mid(vsQuery.EditText, InStr(1, vsQuery.EditText, ":") + 1)), "00")
                        End If
                    End If
                Else
                    vsQuery.EditText = Format(vsQuery.EditText, "0000")
                    vsQuery.EditText = Mid(vsQuery.EditText, 1, 2) & ":" & Mid(vsQuery.EditText, 3)
                End If
        End Select
    End If
    
    If Len(vsQuery.EditText) <> 5 Then
        Cancel = True
        MsgBox "El formato a ingresar es Hora:Minutos (##;## o ####)", vbExclamation, "Atención"
    End If
    
End Sub
Private Sub vsQuery_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error Resume Next
    If NewRow = 0 Then Exit Sub
    vsQuery.CellBorder vbBlack, 2, 2, 2, 2, 0, 0
End Sub

Private Sub vsQuery_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
On Error Resume Next
    vsQuery.CellBorder vbBlack, 0, 0, 0, 0, 0, 0
End Sub

Private Sub OcultoComentarios()
Dim iRow As Integer
    For iRow = 1 To vsQuery.Rows - 1
        If vsQuery.IsSubtotal(iRow) Then
            vsQuery.IsCollapsed(iRow) = flexOutlineCollapsed
        End If
    Next
End Sub


