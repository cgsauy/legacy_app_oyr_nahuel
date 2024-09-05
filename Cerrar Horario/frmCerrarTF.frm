VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCerrarTF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda de Fletes"
   ClientHeight    =   4965
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCerrarTF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   7440
   Begin MSComCtl2.DTPicker dtpCierre 
      Height          =   315
      Left            =   1800
      TabIndex        =   14
      Top             =   1800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   42991617
      CurrentDate     =   38223
   End
   Begin VB.TextBox tZona 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   25
      TabIndex        =   5
      Top             =   1320
      Width           =   3255
   End
   Begin VB.ComboBox cSubGrupo 
      Height          =   315
      Left            =   5400
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox tZonaGrupoZona 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   25
      TabIndex        =   4
      Top             =   1320
      Width           =   3255
   End
   Begin MSComctlLib.TabStrip tsAgenda 
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      MultiRow        =   -1  'True
      Style           =   1
      TabFixedWidth   =   1057
      TabMinWidth     =   1057
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Principal"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Grupo Zona"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Zona"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   635
      ButtonWidth     =   1931
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Editar"
            Key             =   "modificar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "eliminar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
            Key             =   "grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
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
      TabIndex        =   11
      Top             =   4710
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox tDescripcion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   25
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   0
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
            Picture         =   "frmCerrarTF.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCerrarTF.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCerrarTF.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCerrarTF.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCerrarTF.frx":088A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsAgenda 
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3413
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
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
      BackColorFixed  =   -2147483636
      ForeColorFixed  =   -2147483634
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483633
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   8
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
      AutoResize      =   0   'False
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
   Begin VSFlex6DAOCtl.vsFlexGrid vsSubGrupo 
      Height          =   2775
      Left            =   4560
      TabIndex        =   9
      Top             =   1800
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4895
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483636
      ForeColorFixed  =   -2147483634
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
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
      AutoResize      =   0   'False
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Primer día disponible:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   1635
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&SubGrupo:"
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   1320
      Width           =   795
   End
   Begin VB.Label lEs 
      BackStyle       =   0  'Transparent
      Caption         =   "&Grupo Zona:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   915
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Agenda"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   795
   End
   Begin VB.Label lTFNeedAgencia 
      BackStyle       =   0  'Transparent
      Caption         =   "&Flete:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   795
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
Attribute VB_Name = "frmCerrarTF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type tFZona
    Zona As Integer
    ZonNombre As String
    FAbierto As String
    Agenda As Long
    AgendaH As Long
    sAgenda As String
    sOpen As String                 'Este lo uso para indicar cual abro.
End Type
Dim arrFZA() As tFZona
Private douAgenda As Double, douHabilitado As Double
Private dCierre As Date

Private Sub cSubGrupo_Change()
    If cSubGrupo.ListIndex = -1 Then loc_HideShowSG
End Sub

Private Sub cSubGrupo_Click()
    loc_HideShowSG
End Sub

Private Sub dtpCierre_Change()
Dim douZHab As Double, douZAge As Double, dZDate As Date
    
    If dtpCierre.Enabled And dtpCierre.Name = Me.ActiveControl.Name Then
        LimpioDiasHabilitados
        Select Case tsAgenda.SelectedItem.Index
            Case 1: ArmoAgendaFleteEnGrilla douAgenda, douHabilitado, dtpCierre.Value
            Case 2
                loc_SetRangoHsHoraEnvio vsSubGrupo.Cell(flexcpData, vsSubGrupo.Row, 1), douZAge, douZHab, dZDate
                ArmoAgendaFleteEnGrilla douZAge, douZHab, dtpCierre.Value
            Case 3
                loc_SetRangoHsHoraEnvio Val(tZona.Tag), douZAge, douZHab, dZDate
                ArmoAgendaFleteEnGrilla douZAge, douZHab, dtpCierre.Value
        End Select
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    
    ObtengoSeteoForm Me
    frm_OcultoCtrl
    CargoDiasHabilitados
    
    With vsSubGrupo
        .Cols = 1
        .Rows = 1
        .FormatString = "SG|Zona"
        .ExtendLastCol = True
    End With
    Exit Sub
errLoad:
    clsGeneral.OcurrioError "Error al inciar el formulario.", Err.Description
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
End Sub

Private Sub CargoDiasHabilitados()
On Error GoTo errCDH
Dim lngCodigo As Long, intCol As Integer, intDia As Integer
Dim rsHora As rdoResultset
Dim Cons As String

    With vsAgenda
        
        'Cargo las horaenvio que están en horarioflete.
        Cons = "Select HEnNombre, HEnCodigo, HEnIndice, HEnInicio, HEnFin From HoraEnvio " & _
                    " Where HEnCodigo IN (Select Distinct(HFLCodigo) From HorarioFlete) Order By HEnIndice"
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
                
                lngCodigo = rsHora!HEnCodigo: .Cell(flexcpData, 0, .Cols - 1) = lngCodigo
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
                .AddItem DiaSemana(intDia) & Space(5)
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
    clsGeneral.OcurrioError "Error al inicializar la agenda.", Err.Description
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
    Erase arrFZA
    GuardoSeteoForm Me
    cBase.Close
    Set clsGeneral = Nothing
End Sub

Private Sub lTFNeedAgencia_Click()
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
    If Val(tDescripcion.Tag) > 0 Then
        tDescripcion.Tag = ""
        frm_OcultoCtrl
    End If
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
        If Val(tDescripcion.Tag) = 0 Then
            loc_FindByText tDescripcion, "Select TFLCodigo, TFLDescripcion as 'Nombre' From TipoFlete " & _
                                                    "Where TFlDescripcion like '" & Replace(tDescripcion.Text, " ", "%") & "%' Order by 2"
            If Val(tDescripcion.Tag) > 0 Then
                fnc_GetDatosPcpal
                ArmoAgendaFlete
            End If
        End If
        miBotones (Val(tDescripcion.Tag) > 0), False, False
        tsAgenda.Enabled = (Val(tDescripcion.Tag) > 0)
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
    
    With tDescripcion
        .Enabled = Not bolGrabar
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
    End With
    
    vsAgenda.Enabled = bolGrabar
    dtpCierre.Enabled = bolGrabar
    
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
    
    With vsAgenda
        For Fila = 1 To .Rows - 1
            For Col = 1 To .Cols - 1
                If Val(.Cell(flexcpData, Fila, Col)) > 0 Then
                    .Cell(flexcpChecked, Fila, Col) = flexUnchecked
                    .Cell(flexcpBackColor, Fila, Col) = vbButtonFace
                End If
            Next Col
        Next Fila
    End With
    
End Sub

Private Sub AccionModificar()
    
    If tsAgenda.SelectedItem.Index > 1 Then
        If tsAgenda.SelectedItem.Index = 2 And Val(tZonaGrupoZona.Tag) = 0 Then
            MsgBox "Seleccione un grupo zona.", vbExclamation, "Atención"
            tZonaGrupoZona.SetFocus
            Exit Sub
        ElseIf tsAgenda.SelectedItem.Index = 3 And Val(tZona.Tag) = 0 Then
            MsgBox "Seleccione una zona.", vbExclamation, "Atención"
            tZona.SetFocus
            Exit Sub
        ElseIf tsAgenda.SelectedItem.Index = 2 Then
            If vsSubGrupo.Rows = vsSubGrupo.FixedRows Then
                MsgBox "No hay zonas para este grupo.", vbInformation, "Atención"
                Exit Sub
            End If
        End If
    Else
        'Cargo todas las zonas que tenga el tipo de flete.
        ReDim arrFZA(0)
        db_LoadArrayZona
    End If
    tsAgenda.Enabled = False
    cSubGrupo.Enabled = False
    vsSubGrupo.Enabled = False
    tZonaGrupoZona.Enabled = False
    tZona.Enabled = False
    miBotones False, True, True
    vsAgenda.SetFocus
    
End Sub

Private Sub AccionCancelar()
On Error Resume Next
    Erase arrFZA
    miBotones True, False, False
    tDescripcion.SetFocus
    tsAgenda.Enabled = True
    ArmoAgendaFlete
    With tZona
        .Enabled = tsAgenda.SelectedItem.Index = 3
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
    End With
    With tZonaGrupoZona
        .Enabled = tsAgenda.SelectedItem.Index = 2
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
    End With
    With cSubGrupo
        .Enabled = tsAgenda.SelectedItem.Index = 2
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
    End With
    vsSubGrupo.Enabled = (tsAgenda.SelectedItem.Index = 2 And Val(tZonaGrupoZona.Tag) > 0)
    
End Sub

Private Sub ArmoAgendaFlete()
    LimpioDiasHabilitados
    Select Case tsAgenda.SelectedItem.Index
        Case 1: ArmoAgendaFleteEnGrilla douAgenda, douHabilitado, dCierre
        Case 2: ArmoAgendaGrupoZona
        Case 3: ArmoAgendaZona
    End Select
End Sub

Private Sub ArmoAgendaFleteEnGrilla(ByVal douFleteAgenda As Double, ByVal douAgeHab As Double, ByVal dFCierre As Date)
Dim strMat As String, strAux As String
On Error GoTo errSalir

    Screen.MousePointer = 11
    
    If douFleteAgenda = 0 Then douFleteAgenda = douAgenda: douAgeHab = douHabilitado
'    If douAgeHab = 0 Then douAgeHab = douHabilitado
    
    'Armo la grilla en base a la agenda.
    strMat = superp_MatrizSuperposicion(douFleteAgenda)
    If strMat = "" Then GoTo errSalir
    
    Do While strMat <> ""
        If InStr(1, strMat, ",") > 0 Then
            MarcoAgendaEnGrilla CInt(Mid(strMat, 1, InStr(1, strMat, ",") - 1))
            strMat = Mid(strMat, InStr(1, strMat, ",") + 1, Len(strMat))
        Else
            MarcoAgendaEnGrilla CInt(strMat)
            strMat = ""
        End If
    Loop

        'Dada la fecha que tenga como cierre evaluo.
    If DateDiff("d", dFCierre, Date) >= 7 Then
        'Es menor a 1 semana entonces tomo la agenda normal y parto de hoy como fecha de cierre.
        dtpCierre.Value = Date
        OrdenoGrillaPorDia Date
        strMat = superp_MatrizSuperposicion(douFleteAgenda)
    Else
        dtpCierre.Value = dFCierre
        OrdenoGrillaPorDia dFCierre
        strMat = superp_MatrizSuperposicion(douAgeHab)
    End If
    MarcoHabilitadosEnGrilla strMat

errSalir:
    Screen.MousePointer = 0
End Sub

Private Sub MarcoHabilitadosEnGrilla(strMat As String)
    Do While strMat <> ""
        If InStr(1, strMat, ",") > 0 Then
            MarcoEnGrilla CInt(Mid(strMat, 1, InStr(1, strMat, ",") - 1))
            strMat = Mid(strMat, InStr(1, strMat, ",") + 1, Len(strMat))
        Else
            MarcoEnGrilla CInt(strMat)
            strMat = ""
        End If
    Loop
End Sub

Private Sub OrdenoGrillaPorDia(dFecha As Date)
Dim intDiaSemana As Integer, I As Integer
Dim intCant As Integer, intA As Integer, intPos As Integer

    intDiaSemana = Weekday(dFecha)
    
    With vsAgenda
        intPos = 1
        '--------------------------------------------------------------------
        For I = 1 To 7
            intA = UbicacionDiaSemana(intDiaSemana)
            If intA > 0 Then
                .RowPosition(intA) = intPos
                .Cell(flexcpText, intPos, 0) = DiaSemana(intDiaSemana) & " " & Day(dFecha + intPos - 1)
                intPos = intPos + 1
            End If
            If intDiaSemana = 7 Then intDiaSemana = 1 Else intDiaSemana = intDiaSemana + 1
        Next I
        '--------------------------------------------------------------------
    End With
    
End Sub

Private Function UbicacionDiaSemana(intDia As Integer) As Integer
Dim j As Integer
    
    UbicacionDiaSemana = 0
    For j = 1 To vsAgenda.Rows - 1
        'Busco si tengo el día que busco
        If intDia = Val(vsAgenda.Cell(flexcpData, j, 0)) Then UbicacionDiaSemana = j: Exit Function
    Next j
    
End Function

Private Sub MarcoEnGrilla(intIndice As Integer)
On Error Resume Next
Dim Fila As Integer, Col As Integer
    For Fila = 1 To vsAgenda.Rows - 1
        For Col = 1 To vsAgenda.Cols - 1
            If Val(vsAgenda.Cell(flexcpData, Fila, Col)) = intIndice Then vsAgenda.Cell(flexcpChecked, Fila, Col) = flexChecked
        Next Col
    Next Fila
End Sub

Private Sub MarcoAgendaEnGrilla(intIndice As Integer)
Dim Fila As Integer, Col As Integer
    
    For Fila = 1 To vsAgenda.Rows - 1
        For Col = 1 To vsAgenda.Cols - 1
            If Val(vsAgenda.Cell(flexcpData, Fila, Col)) = intIndice Then
                vsAgenda.Cell(flexcpBackColor, Fila, Col) = vbWindowBackground
            End If
        Next Col
    Next Fila
    
End Sub


Private Sub AccionGrabar()
Dim douAux As Double
    
    If MsgBox("¿Confirma almacenar la agenda ingresada?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
        On Error GoTo errGrabar
        Screen.MousePointer = 11
        douAux = CalculoValorSuperposicion
        If tsAgenda.SelectedItem.Index = 1 Then
            db_SavePrincipal douAux
            dCierre = dtpCierre.Value
        ElseIf tsAgenda.SelectedItem.Index = 2 Then
            'Para todas las zonas del grupo armo la agenda.
            If Not db_SaveGrupoZona(douAux) Then Exit Sub
        Else
            'Zona
            db_SaveAgendaZona Val(tZona.Tag), douAux
        End If
        AccionCancelar
        Screen.MousePointer = 0
    End If
    Exit Sub
errGrabar:
    clsGeneral.OcurrioError "Ocurrió un error al intentar almacenar la información.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Function GetIndexClose(ByVal sMatNormal As String, sMatAbierta As String) As String
Dim arrN() As String
Dim iQ As Integer
Dim sRet As String
    If sMatNormal = "" Then Exit Function
    arrN = Split(sMatNormal, ",")
    For iQ = 0 To UBound(arrN)
        If InStr(1, "," & sMatAbierta & ",", "," & arrN(iQ) & ",") = 0 Then
            If sRet <> "" Then sRet = sRet & ","
            sRet = sRet & arrN(iQ)
        End If
    Next
    GetIndexClose = sRet
End Function
Private Sub db_SavePrincipal(ByVal douAux As Double)
Dim Cons As String
Dim rsAux As rdoResultset
Dim sMat As String
Dim arrClose() As String, arrAux() As String
       
    If UBound(arrFZA) > 0 Then
        'Dada la matriz del flete busco los índices que están cerrados.
        sMat = GetIndexClose(superp_MatrizSuperposicion(douAgenda), superp_MatrizSuperposicion(douAux))
        'Array de días (índices) que están cerrados y están en la agenda.
        arrClose = Split(sMat, ",")
    End If
        
    Cons = "Select * from TipoFlete Where TFlCodigo = " & Val(tDescripcion.Tag)
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    rsAux.Edit
    rsAux!TFlAgendaHabilitada = douAux
    rsAux!TFLFechaAgeHab = Format(dtpCierre.Value, "mm/dd/yyyy 00:00:00")
    rsAux.Update
    rsAux.Close
    douHabilitado = douAux
    
    Dim sMH As String
    Dim iQ As Integer, iQ2 As Integer, iQ3 As Integer
    Dim arrOpen() As String
    
    'Tengo que buscar las zonas que afecta el cambio y no las modifique antes.
    'Este paso lo hago de la sgte forma, recorro uno a uno los que no afecto el cambio.
    For iQ = 1 To UBound(arrFZA)
        Cons = "Select * From FleteAgendaZona Where FAZTipoFlete = " & Val(tDescripcion.Tag) & " And FAZZona =  " & arrFZA(iQ).Zona
        Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            
            'Si la fecha es mayor a la que tiene la zona considero esta fecha.
            sMat = superp_MatrizSuperposicion(rsAux!FAZAgenda)
            If rsAux!FAZFechaAgeHab >= DateAdd("d", -7, Date) And rsAux!FAZFechaAgeHab > dtpCierre.Value Then
                sMH = superp_MatrizSuperposicion(rsAux!FAZAgendaHabilitada)
            Else
                sMH = sMat
            End If
            
            If arrFZA(iQ).sOpen <> "" Then
                If InStr(1, arrFZA(iQ).sOpen, ",") > 0 Then
                    arrOpen = Split(arrFZA(iQ).sOpen, ",")
                Else
                    ReDim arrOpen(0)
                    arrOpen(0) = arrFZA(iQ).sOpen
                End If
                
                For iQ2 = 0 To UBound(arrOpen)
                    If InStr(1, "," & sMat & ",", "," & arrOpen(iQ2) & ",") > 0 And InStr(1, "," & sMH & ",", "," & arrOpen(iQ2) & ",") = 0 Then
                        If sMH <> "" Then sMH = sMH & ", "
                        sMH = sMH & arrOpen(iQ2)
                    End If
                Next
                
            End If
            
            Erase arrOpen
            'Dada la agenda habilitada que queda.
            For iQ2 = 0 To UBound(arrClose)
                If IsNumeric(arrClose(iQ2)) Then
                    If InStr(1, "," & sMH & ",", "," & arrClose(iQ2) & ",") > 0 Then
                        If InStr(1, sMH, ",") > 0 Then
                            arrOpen = Split(sMH, ",")
                        Else
                            ReDim arrOpen(0)
                            arrOpen(0) = sMH
                        End If
                        sMH = ""
                        For iQ3 = 0 To UBound(arrOpen)
                            If arrOpen(iQ3) <> arrClose(iQ2) Then
                                If sMH <> "" Then sMH = sMH & ","
                                sMH = sMH & arrOpen(iQ3)
                            End If
                        Next
                    End If
                End If
            Next iQ2
            
            Dim douA As Double
            
            Erase arrAux
            If InStr(1, sMH, ",") > 0 Then
                arrAux = Split(sMH, ",")
            Else
                ReDim arrAux(0)
                arrAux(0) = sMH
            End If
            douA = 0
            For iQ2 = 0 To UBound(arrAux)
                If IsNumeric(arrAux(iQ2)) Then douA = douA + superp_ValSuperposicion(CInt(arrAux(iQ2)))
            Next
            
            rsAux.Edit
            rsAux!FAZAgendaHabilitada = douA
            If dtpCierre.Value > rsAux!FAZFechaAgeHab Then rsAux!FAZFechaAgeHab = Format(dtpCierre.Value, "mm/dd/yyyy 00:00:00")
            rsAux.Update
            
        End If
        rsAux.Close
    Next
    
End Sub
Private Function db_SaveGrupoZona(ByVal douAge As Double) As Boolean
Dim iQ As Integer, iSG As Integer
    
    db_SaveGrupoZona = False
    On Error GoTo errBT
    cBase.BeginTrans
    On Error GoTo errRB
    
    With vsSubGrupo
        iSG = .Cell(flexcpValue, .Row, 0)
        For iQ = .FixedRows To .Rows - 1
            If .Cell(flexcpValue, iQ, 0) = iSG Then
                db_SaveAgendaZona .Cell(flexcpData, iQ, 1), douAge
            End If
        Next
    End With
    cBase.CommitTrans
    db_SaveGrupoZona = True
    Exit Function
    
errBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al iniciar la transacción.", Err.Description
    Exit Function
errResumo:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al intentar grabar la información.", Err.Description
    Exit Function
errRB:
    Resume errRB
End Function

Private Sub db_SaveAgendaZona(ByVal lCodZona As Long, ByVal douAgen As Double)
Dim Cons As String
Dim rsAux As rdoResultset
    
    Cons = "Select * From FleteAgendaZona " & _
                " Where FAZZona = " & lCodZona & " And FAZTipoFlete = " & Val(tDescripcion.Tag)
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If rsAux.EOF Then
        rsAux.AddNew
        rsAux!FAZZona = lCodZona
        rsAux!FAZTipoFlete = Val(tDescripcion.Tag)
        rsAux!FAZAgenda = douAgen
    Else
        rsAux.Edit
    End If
    rsAux!FAZAgendaHabilitada = douAgen
    rsAux!FAZFechaAgeHab = Format(dtpCierre.Value, "mm/dd/yyyy 00:00:00")
    rsAux.Update
    rsAux.Close
    
End Sub

Private Sub tsAgenda_Click()
    
    With tZonaGrupoZona
        .Enabled = (tsAgenda.SelectedItem.Index = 2)
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
        .Visible = (tsAgenda.SelectedItem.Index <> 3)
    End With
    With tZona
        .Enabled = (tsAgenda.SelectedItem.Index = 3)
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
        .Visible = (tsAgenda.SelectedItem.Index = 3)
    End With
    With cSubGrupo
        .Enabled = tsAgenda.SelectedItem.Index = 2
        .Tag = "": .Clear
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
    End With
    vsSubGrupo.Rows = 1
    vsSubGrupo.Enabled = cSubGrupo.Enabled
    
    lEs.Caption = IIf(tsAgenda.SelectedItem.Index = 3, "&Zona:", "&Grupo Zona:")
    
    If Val(tDescripcion.Tag) > 0 And tsAgenda.Enabled Then
        ArmoAgendaFlete
    End If
    
End Sub

Private Sub tZona_Change()
    If Val(tZona.Tag) > 0 Then
        tZona.Tag = ""
        LimpioDiasHabilitados
    End If
End Sub

Private Sub tZona_GotFocus()
    With tZona
        If .Enabled Then .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tZona_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn And Trim(tZona.Text) <> "" Then
        If Val(tZona.Tag) = 0 Then
            loc_FindByText tZona, "Select ZonCodigo, ZonNombre as 'Nombre' From Zona Where ZonNombre like '" & Replace(tZona.Text, " ", "%") & "%' Order by ZonNombre"
            If Val(tZona.Tag) > 0 Then ArmoAgendaFlete
        End If
    End If
    
End Sub

Private Sub tZonaGrupoZona_Change()
    If Val(tZonaGrupoZona.Tag) > 0 Then
        tZonaGrupoZona.Tag = ""
        LimpioDiasHabilitados
        If cSubGrupo.Enabled Then cSubGrupo.Clear: cSubGrupo.Clear: vsSubGrupo.Rows = 1
    End If
End Sub

Private Sub tZonaGrupoZona_GotFocus()
    With tZonaGrupoZona
        If .Enabled Then .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tZonaGrupoZona_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(tZonaGrupoZona.Text) <> "" Then
        If Val(tZonaGrupoZona.Tag) > 0 Then
            If cSubGrupo.Enabled Then cSubGrupo.SetFocus
        Else
            loc_FindByText tZonaGrupoZona, "Select GZoCodigo, GZoNombre as 'Nombre' From GrupoZona Where GZoNombre like '" & Replace(tZonaGrupoZona.Text, " ", "%") & "%' Order by GZoNombre"
            If Val(tZonaGrupoZona.Tag) > 0 Then ArmoAgendaFlete
        End If
    End If
End Sub

Private Sub vsAgenda_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
    If Val(vsAgenda.Cell(flexcpData, Row, Col)) = 0 Then Cancel = True
End Sub

Private Sub vsAgenda_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim iQ As Integer, iQ2 As Integer
Dim arrV() As String

If tsAgenda.SelectedItem.Index > 1 Then Exit Sub

    If vsAgenda.Cell(flexcpChecked, Row, Col) <> flexChecked Then
        Dim sMat As String, sMat2 As String
        'Esta desmarcando x lo tanto tengo que ver si esta en otras zonas.
        If DateDiff("d", dCierre, Date) >= 7 Then
            sMat = superp_MatrizSuperposicion(douAgenda)
        Else
            sMat = superp_MatrizSuperposicion(douHabilitado)
        End If
        'Aca tengo el valor de la agenda con la que inicio.
'        If InStr(1, "," & sMat & ",", "," & vsAgenda.Cell(flexcpData, Row, Col) & ",") = 0 Then
            'Este valor estaba marcado x lo tanto recorro las zonas que tiene este valor en la agenda habilitada.
            For iQ = 1 To UBound(arrFZA)
                If CDate(arrFZA(iQ).FAbierto) <= Date Then
                    If DateDiff("d", CDate(arrFZA(iQ).FAbierto), Date) < 7 And DateDiff("d", CDate(arrFZA(iQ).FAbierto), Date) >= 0 Then
                        sMat2 = superp_MatrizSuperposicion(arrFZA(iQ).AgendaH)
                    Else
                        sMat2 = arrFZA(iQ).sAgenda
                    End If
                    If InStr(1, "," & sMat2 & ",", "," & vsAgenda.Cell(flexcpData, Row, Col) & ",") = 0 _
                        And InStr(1, "," & arrFZA(iQ).sAgenda & ",", "," & vsAgenda.Cell(flexcpData, Row, Col) & ",") > 0 Then
                        If MsgBox("¿La zona """ & arrFZA(iQ).ZonNombre & """ tiene el día cerrado." & vbCr & vbCr & "¿Desea abrirlo?", vbQuestion + vbYesNo, "Día cerrado") = vbYes Then
                            If arrFZA(iQ).sOpen <> "" Then arrFZA(iQ).sOpen = arrFZA(iQ).sOpen & ","
                            arrFZA(iQ).sOpen = arrFZA(iQ).sOpen & vsAgenda.Cell(flexcpData, Row, Col)
                            vsAgenda.Cell(flexcpBackColor, Row, Col) = &HC0E0FF
                        End If
                    End If
                End If
            Next
'        End If
    Else
        vsAgenda.Cell(flexcpBackColor, Row, Col) = vbWindowBackground
        For iQ = 1 To UBound(arrFZA)
            Erase arrV
            If InStr(1, "," & arrFZA(iQ).sOpen & ",", "," & vsAgenda.Cell(flexcpData, Row, Col) & ",") > 0 Then
                If InStr(arrFZA(iQ).sOpen, ",") > 0 Then
                    arrV = Split(arrFZA(iQ).sOpen, ",")
                    arrFZA(iQ).sOpen = ""
                    For iQ2 = 0 To UBound(arrV)
                        If arrV(iQ2) <> vsAgenda.Cell(flexcpData, Row, Col) Then
                            If arrFZA(iQ).sOpen <> "" Then arrFZA(iQ).sOpen = arrFZA(iQ).sOpen & ","
                            arrFZA(iQ).sOpen = arrFZA(iQ).sOpen & arrV(iQ2)
                        End If
                    Next iQ2
                Else
                    arrFZA(iQ).sOpen = ""
                End If
            End If
        Next iQ
    End If
    
End Sub

Private Sub ArmoAgendaGrupoZona()
Dim Cons As String
Dim rsAux As rdoResultset
Dim sZonaIn As String
Dim douAgendaAux As Double
Dim lColor As Long, lAux As Long
Dim iSG As Integer
    
    With cSubGrupo
        .Clear
        .AddItem "Sub Grupo 0"
        .ItemData(.NewIndex) = 0
    End With
    vsSubGrupo.Rows = 1
    If Val(tZonaGrupoZona.Tag) = 0 Then Exit Sub
    iSG = 0
    
    'Paso 1 Cargo todos los subgrupos que pueda formar en base a la tabla FleteAgendaZona
    'Paso 2 Cargo el resto de las zonas y las asigno como GrupoPcpal al primero.
    sZonaIn = "0"
    douAgendaAux = -1
    lColor = &HEEFFFF
    'Si la agenda es nula se considera la agenda de la pcpal.
    Cons = "Select FAZZona, ZonNombre, IsNull(FAZAgenda, 0)  as Agenda, FAZAgendaHabilitada, FAZFechaAgeHab From FleteAgendaZona, Zona " & _
        " Where FAZTipoFlete = " & Val(tDescripcion.Tag) & _
        " And FAZZona IN (Select GZZZona From GrupoZonaZona Where GZZGrupo = " & Val(tZonaGrupoZona.Tag) & ")" & _
        " And FAZZona = ZonCodigo Order By FAZAgenda"
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        If douAgendaAux <> rsAux!Agenda Then
            lColor = IIf(lColor = &HEEFFFF, vbWhite, &HEEFFFF)
            If rsAux!Agenda > 0 Then iSG = iSG + 1
            If iSG > 0 Then
                With cSubGrupo
                    .AddItem "Sub Grupo " & iSG
                    .ItemData(.NewIndex) = iSG
                End With
            End If
            douAgendaAux = rsAux!Agenda
        End If
        With vsSubGrupo
            .AddItem iSG
            .Cell(flexcpText, .Rows - 1, 1) = Trim(rsAux!ZonNombre)
            douAgendaAux = rsAux!Agenda
            .Cell(flexcpData, .Rows - 1, 0) = douAgendaAux
            lAux = rsAux!FAZZona
            .Cell(flexcpData, .Rows - 1, 1) = lAux
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = lColor
        End With
        sZonaIn = sZonaIn & ", " & rsAux!FAZZona
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    If vsSubGrupo.Rows > vsSubGrupo.FixedRows Then
        If vsSubGrupo.Cell(flexcpValue, 1, 0) = 0 Then
            lColor = vsSubGrupo.Cell(flexcpBackColor, 1, 0)
        Else
            'La primera si ó si fue blanca
            lColor = &HEEFFFF
        End If
    Else
        lColor = vbWhite
    End If
        
    'Primero cargo todos los GrupoZona que tengo para el tipoflete
    Cons = "Select Distinct(ZonCodigo), ZonNombre From GrupoZonaZona, Zona" & _
            " Where GZZGrupo = " & Val(tZonaGrupoZona.Tag) & _
            " And GZZZona Not IN(" & sZonaIn & ")" & _
            " And GZZZona = ZonCodigo Order by ZonNombre"
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        With vsSubGrupo
            .AddItem 0, .FixedRows
            .Cell(flexcpText, .FixedRows, 1) = Trim(rsAux!ZonNombre)
            .Cell(flexcpData, .FixedRows, 0) = 0
            lAux = rsAux!ZonCodigo
            .Cell(flexcpData, .FixedRows, 1) = lAux
            .Cell(flexcpBackColor, .FixedRows, 0, , .Cols - 1) = lColor
        End With
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    With vsSubGrupo
        If .Rows > .FixedRows Then
            .Select .FixedRows, 0, .Rows - 1, .Cols - 1
            .Sort = flexSortStringAscending
            .Select .FixedRows, 0, .FixedRows, 0
            Call vsSubGrupo_RowColChange
        End If
    End With
    
End Sub

Private Sub loc_FindByText(txtCtrl As TextBox, ByVal Cons As String)
On Error GoTo errCTF
Dim rsAux As rdoResultset
Dim sNombre As String, lCodigo As Long
    
    Screen.MousePointer = 11
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If rsAux.EOF Then
        rsAux.Close
        MsgBox "No se encontraron datos para el filtro ingresado.", vbExclamation, "Atención"
    Else
        rsAux.MoveNext
        If Not rsAux.EOF Then
            rsAux.Close
            lCodigo = fnc_HelpList(Cons, "Grupos de Zona", sNombre)
        Else
            rsAux.MoveFirst
            sNombre = Trim(rsAux(1))
            lCodigo = rsAux(0)
            rsAux.Close
        End If
    End If
    If lCodigo > 0 Then
        With txtCtrl
            .Text = sNombre
            .Tag = lCodigo
        End With
    End If
    Screen.MousePointer = 0
    Exit Sub
errCTF:
    clsGeneral.OcurrioError "Error al buscar.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function fnc_HelpList(ByVal Cons As String, sTitulo As String, sRetNombre As String) As Long
    
    'Valores a retornar código y nombre
    fnc_HelpList = 0
    sRetNombre = ""
    
    Dim objLista As New clsListadeAyuda
    If objLista.ActivarAyuda(cBase, Cons, 4500, 1, sTitulo) > 0 Then
        fnc_HelpList = objLista.RetornoDatoSeleccionado(0)
        sRetNombre = objLista.RetornoDatoSeleccionado(1)
    End If
    Set objLista = Nothing

End Function

Private Sub fnc_GetDatosPcpal()
On Error GoTo errGDP
Dim Cons As String
Dim rsAux As rdoResultset
    
    douAgenda = 0
    douHabilitado = 0
    dCierre = Date
    Cons = "Select IsNull(TFLAgenda, 0), TFLAgendaHabilitada,  TFLFechaAgeHab From TipoFlete Where TFLCodigo = " & Val(tDescripcion.Tag)
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        douAgenda = rsAux(0)
        If Not IsNull(rsAux!TFlAgendaHabilitada) Then douHabilitado = rsAux!TFlAgendaHabilitada Else douHabilitado = douAgenda
        If Not IsNull(rsAux!TFLFechaAgeHab) Then dCierre = rsAux!TFLFechaAgeHab
    End If
    rsAux.Close
    Exit Sub
errGDP:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar la agenda principal del tipo de flete.", Err.Description
End Sub

Private Sub frm_OcultoCtrl()
    lTFNeedAgencia.Tag = ""
    LimpioDiasHabilitados
    miBotones False, False, False
    douAgenda = 0
    douHabilitado = 0
    dCierre = Date
    tsAgenda.Tabs(1).Selected = True
    tsAgenda_Click
    tsAgenda.Enabled = False
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


Private Sub vsSubGrupo_DblClick()
Dim lZona As Long, sNombre As String

    If vsSubGrupo.Row = 0 Then Exit Sub
    
    lZona = vsSubGrupo.Cell(flexcpData, vsSubGrupo.Row, 1)
    sNombre = vsSubGrupo.Cell(flexcpText, vsSubGrupo.Row, 1)
    
    'Considero edición de zona.
    tsAgenda.Tabs(3).Selected = True
    With tZona
        .Text = sNombre
        .Tag = lZona
    End With
    ArmoAgendaFlete
    AccionModificar
    
End Sub

Private Sub vsSubGrupo_RowColChange()
Dim douZHab As Double, douZAge As Double, dZDate As Date
    If vsSubGrupo.Row = 0 Then Exit Sub
    LimpioDiasHabilitados
    loc_SetRangoHsHoraEnvio vsSubGrupo.Cell(flexcpData, vsSubGrupo.Row, 1), douZAge, douZHab, dZDate
    If DateDiff("d", dZDate, Date) >= 7 Then
        dtpCierre.Value = Date
    Else
        dtpCierre.Value = dZDate
    End If
    ArmoAgendaFleteEnGrilla douZAge, douZHab, dZDate
End Sub

Private Sub ArmoAgendaZona()
Dim douAux As Double, dFecha As Date, douHab As Double

    If Val(tZona.Tag) = 0 Then Exit Sub
    loc_SetRangoHsHoraEnvio Val(tZona.Tag), douAux, douHab, dFecha
    ArmoAgendaFleteEnGrilla douAux, douHab, dFecha
    
End Sub

Private Sub loc_HideShowSG()
Dim iQ As Integer, iQSel As Integer
Dim iSel As Integer

    LimpioDiasHabilitados
    iQSel = 0
    If cSubGrupo.ListIndex = -1 Then
        iSel = -1
    Else
        iSel = cSubGrupo.ItemData(cSubGrupo.ListIndex)
    End If
    With vsSubGrupo
        For iQ = .FixedRows To .Rows - 1
            If iSel = -1 Then
                .RowHidden(iQ) = False
            Else
                .RowHidden(iQ) = Not (.Cell(flexcpValue, iQ, 0) = iSel)
            End If
            If Not .RowHidden(iQ) Then iQSel = iQ
        Next
    End With
    If iQSel > 0 Then vsSubGrupo.Select iQSel, 0
End Sub

Private Sub loc_SetRangoHsHoraEnvio(ByVal lZona As Long, ByRef douAge As Double, ByRef douHab As Double, ByRef dDate As Date)
On Error GoTo errSRH
Dim Cons As String
Dim rsAux As rdoResultset
    
    douAge = douAgenda
    douHab = douHabilitado
    dDate = dCierre
    
    Cons = "Select FAZAgenda, FAZAgendaHabilitada, FAZFechaAgeHAb From FleteAgendaZona " & _
        " Where FAZTipoFlete = " & Val(tDescripcion.Tag) & _
        " And FAZZona = " & lZona
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux!FAZAgenda) Then
            If Not IsNull(rsAux!FAZAgenda) Then douAge = rsAux!FAZAgenda Else douAge = douAgenda
            If Not IsNull(rsAux!FAZAgendaHabilitada) Then douHab = rsAux!FAZAgendaHabilitada Else douHab = douAge
            If Not IsNull(rsAux!FAZFechaAgeHab) Then dDate = rsAux!FAZFechaAgeHab Else dDate = Date
        End If
    End If
    rsAux.Close
    Exit Sub
errSRH:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar la agenda para la zona.", Err.Description
End Sub

Private Sub db_LoadArrayZona()
Dim Cons As String
Dim rsAux As rdoResultset
Dim sMH As String

    Cons = "Select  FleteAgendaZona.*, rTrim(ZonNombre) as ZN From FleteAgendaZona, Zona " & _
            " Where FAZTipoFlete = " & Val(tDescripcion.Tag) & " And FAZAgenda Is Not Null" & _
            " And FAZZona = ZonCodigo"
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        ReDim Preserve arrFZA(UBound(arrFZA) + 1)
        With arrFZA(UBound(arrFZA))
            .Zona = rsAux!FAZZona
            .ZonNombre = rsAux!ZN
            .Agenda = rsAux!FAZAgenda
            If Not IsNull(rsAux!FAZAgendaHabilitada) Then .AgendaH = rsAux!FAZAgendaHabilitada Else .AgendaH = .Agenda
            If IsNull(rsAux!FAZFechaAgeHab) Then
                .FAbierto = dCierre
                .AgendaH = .Agenda
            Else
                .FAbierto = rsAux!FAZFechaAgeHab
            End If
            .sAgenda = superp_MatrizSuperposicion(.Agenda)
        End With
        rsAux.MoveNext
    Loop
    rsAux.Close
    
End Sub
