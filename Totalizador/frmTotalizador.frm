VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{191D08B9-4E92-4372-BF17-417911F14390}#1.5#0"; "orGridPreview.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTotalizador 
   Caption         =   "Totalizador de Operaciones"
   ClientHeight    =   6165
   ClientLeft      =   2535
   ClientTop       =   2940
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTotalizador.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   7680
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   60
      ScaleHeight     =   435
      ScaleWidth      =   6195
      TabIndex        =   6
      Top             =   5040
      Width           =   6255
      Begin VB.CommandButton bPreview 
         Height          =   310
         Left            =   1620
         Picture         =   "frmTotalizador.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmTotalizador.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   2340
         Picture         =   "frmTotalizador.frx":0846
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   540
         Picture         =   "frmTotalizador.frx":0948
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   1260
         Picture         =   "frmTotalizador.frx":0D0E
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
   End
   Begin orGridPreview.GridPreview orPrev 
      Left            =   180
      Top             =   4080
      _ExtentX        =   873
      _ExtentY        =   873
      PageBorder      =   3
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      BeginProperty FooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros de consulta"
      ForeColor       =   &H00800000&
      Height          =   630
      Left            =   120
      TabIndex        =   2
      Top             =   30
      Width           =   7275
      Begin AACombo99.AACombo cMonedaN 
         Height          =   315
         Left            =   6180
         TabIndex        =   5
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
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
      Begin AACombo99.AACombo cMonedaT 
         Height          =   315
         Left            =   1980
         TabIndex        =   4
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
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
      Begin VB.CheckBox cConvertir 
         Caption         =   "Convertir moneda extranjera a..."
         Height          =   255
         Left            =   3360
         TabIndex        =   1
         Top             =   280
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Totalizar operaciones en:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   280
         Width           =   1935
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   1500
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   3413
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   5
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
   Begin MSComctlLib.ImageList img1 
      Left            =   6000
      Top             =   2460
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTotalizador.frx":0E10
            Key             =   "salir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTotalizador.frx":112A
            Key             =   "help"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTotalizador.frx":123C
            Key             =   "consultar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTotalizador.frx":1396
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTotalizador.frx":14F0
            Key             =   "print"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTotalizador.frx":164A
            Key             =   "limpiar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTotalizador.frx":175C
            Key             =   "vista2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTotalizador.frx":18B6
            Key             =   "vista1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTotalizador.frx":1A10
            Key             =   "move2"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTotalizador.frx":1B6A
            Key             =   "move3"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTotalizador.frx":1CC4
            Key             =   "move4"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTotalizador.frx":1E1E
            Key             =   "move1"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTotalizador.frx":1F78
            Key             =   "configprint"
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuOpConsultar 
         Caption         =   "&Consultar"
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuOpImprimir 
         Caption         =   "&Imprimr"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuSalFormulario 
         Caption         =   "&Del Formulario"
         Shortcut        =   ^X
      End
      Begin VB.Menu MnuSL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCambiarBase 
         Caption         =   "Cambiar BD"
      End
   End
End
Attribute VB_Name = "frmTotalizador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim aCantidadT As Long, aImporteT As Currency

Dim aTC As Currency
Dim mMoneda As Long

Private Sub AccionConsultar()

    On Error GoTo ErrBC
    If Not ValidoDatos Then Screen.MousePointer = vbDefault: Exit Sub
    
    Screen.MousePointer = 11
    mMoneda = cMonedaT.ItemData(cMonedaT.ListIndex)
    
    If cMonedaN.Enabled Then
        aTC = TasadeCambio(CLng(mMoneda), CLng(cMonedaN.ItemData(cMonedaN.ListIndex)), Date)
    End If
    
    vsLista.Rows = 1
    aCantidadT = 0: aImporteT = 0
            
    cBase.QueryTimeout = 60
    
    'TOTALIZO CGSA----------------------------------------------------------------------------------------------------
    zAddTitulo "CGSA"
    
    'Tipos de Operaciones   a) Normales y En Gestor     b) A Perdida
    If Not CargoCreditosNormalesYGestor Then GoTo etFin
    If Not CargoCreditosAPerdida Then GoTo etFin
    
    zAddTotal "CGSA"
    
    'TOTALIZO MEGA----------------------------------------------------------------------------------------------------
    vsLista.AddItem "": vsLista.AddItem "": vsLista.AddItem "": vsLista.AddItem ""
     aCantidadT = 0: aImporteT = 0
     zAddTitulo "MEGA"
     
    'Tipos de Operaciones   a) Normales y En Gestor     b) A Perdida
    If Not CargoCreditosNormalesYGestor(False) Then GoTo etFin
    If Not CargoCreditosAPerdida(False) Then GoTo etFin
    
    zAddTotal "MEGA"

etFin:
    vsLista.AutoSize 0
   
    cBase.QueryTimeout = 15
    Screen.MousePointer = 0
    Exit Sub

ErrBC:
    clsGeneral.OcurrioError "Error al realizar la consulta." & Trim(Err.Description)
    cBase.QueryTimeout = 15
    Screen.MousePointer = 0
End Sub

Private Function CargoCreditosNormalesYGestor(Optional CGSA As Boolean = True) As Boolean

Dim aCantidad As Long, aImporte As Currency
Dim aCantidadP As Long, aImporteP As Currency
    
    CargoCreditosNormalesYGestor = False
    aCantidadP = 0: aImporteP = 0
    
    With vsLista
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = "Tipo de Operación"
        .Cell(flexcpText, .Rows - 1, 1) = "Q Op."
        .Cell(flexcpText, .Rows - 1, 2) = "Saldo"
        .Cell(flexcpFontItalic, .Rows - 1, 0, , .Cols - 1) = True
        .Cell(flexcpFontUnderline, .Rows - 1, 0, , .Cols - 1) = True
    End With
    
    'Creditos Normales---------------------------------------------------------------------------------------------
    Cons = ArmoConsulta(TipoCredito.Normal, mMoneda, CGSA)

    On Error GoTo errTOutN

eqSQLNormal:

    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    aCantidad = 0: aImporte = 0
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!Cantidad) Then aCantidad = RsAux!Cantidad
        If Not IsNull(RsAux!Importe) Then aImporte = RsAux!Importe
    End If
    RsAux.Close
    
    With vsLista
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = "Operaciones Normales"
        .Cell(flexcpText, .Rows - 1, 1) = Format(aCantidad, "#,##0")
        .Cell(flexcpText, .Rows - 1, 2) = Format(aImporte, "#,##0.00")
    
        If cMonedaN.Enabled Then
                .Cell(flexcpText, .Rows - 1, 3) = Format(aImporte * aTC, "#,##0.00")
        End If
    End With
    
    aImporteP = aImporteP + aImporte
    aCantidadP = aCantidadP + aCantidad
    '------------------------------------------------------------------------------------------------------------------
     
    'Creditos en Gestor---------------------------------------------------------------------------------------------
    Cons = ArmoConsulta(TipoCredito.Gestor, mMoneda, CGSA)

    On Error GoTo errTOutG
eqSQLGestor:
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    aCantidad = 0: aImporte = 0
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!Cantidad) Then aCantidad = RsAux!Cantidad
        If Not IsNull(RsAux!Importe) Then aImporte = RsAux!Importe
    End If
    RsAux.Close
    
    With vsLista
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = "Operaciones en Gestor"
        .Cell(flexcpText, .Rows - 1, 1) = Format(aCantidad, "#,##0")
        .Cell(flexcpText, .Rows - 1, 2) = Format(aImporte, "#,##0.00")
    
        If cMonedaN.Enabled Then
                .Cell(flexcpText, .Rows - 1, 3) = Format(aImporte * aTC, "#,##0.00")
        End If
    End With
    
    aImporteP = aImporteP + aImporte
    aCantidadP = aCantidadP + aCantidad
    '------------------------------------------------------------------------------------------------------------------
    
    'Total de las operaciones
    With vsLista
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = "Total Normales y Gestor"
        .Cell(flexcpText, .Rows - 1, 1) = Format(aCantidadP, "#,##0")
        .Cell(flexcpText, .Rows - 1, 2) = Format(aImporteP, "#,##0.00")
    
        If cMonedaN.Enabled Then .Cell(flexcpText, .Rows - 1, 3) = Format(aImporteP * aTC, "#,##0.00")
        
        .Cell(flexcpFontBold, .Rows - 1, 0, , .Cols - 1) = True
        .AddItem ""
    End With
     
    aImporteT = aImporteT + aImporteP
    aCantidadT = aCantidadT + aCantidadP
    
    CargoCreditosNormalesYGestor = True
    Exit Function
    
errTOutN:
    frmTOut.Show vbModal, Me
    Me.Refresh
    If frmTOut.prmOK Then Resume eqSQLNormal
    Exit Function

errTOutG:
    frmTOut.Show vbModal, Me
    Me.Refresh
    If frmTOut.prmOK Then Resume eqSQLGestor
    Exit Function
End Function

Private Function CargoCreditosAPerdida(Optional CGSA As Boolean = True) As Boolean
    
Dim aCantidad As Long, aImporte As Currency
    
    CargoCreditosAPerdida = False
    
    'Creditos A Perdida---------------------------------------------------------------------------------------------
    Cons = ArmoConsulta(TipoCredito.Incobrable, mMoneda, CGSA)
    
    On Error GoTo errTOutN
    
eqSQL:
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    aCantidad = 0: aImporte = 0
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!Cantidad) Then aCantidad = RsAux!Cantidad
        If Not IsNull(RsAux!Importe) Then aImporte = RsAux!Importe
    End If
    RsAux.Close
    
    With vsLista
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = "Operaciones a Pérdida"
        .Cell(flexcpText, .Rows - 1, 1) = Format(aCantidad, "#,##0")
        .Cell(flexcpText, .Rows - 1, 2) = Format(aImporte, "#,##0.00")
    
        If cMonedaN.Enabled Then .Cell(flexcpText, .Rows - 1, 3) = Format(aImporte * aTC, "#,##0.00")
        
        .AddItem ""
    End With
    
    aImporteT = aImporteT + aImporte
    aCantidadT = aCantidadT + aCantidad
    
    CargoCreditosAPerdida = True
    Exit Function
    
errTOutN:
    frmTOut.Show vbModal, Me
    Me.Refresh
    If frmTOut.prmOK Then Resume eqSQL
    Exit Function
End Function

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub

Private Sub bImprimir_Click()
    AccionImprimir ToPrinter:=True
End Sub

Private Sub bNoFiltros_Click()
    cConvertir.Value = vbUnchecked
    BuscoCodigoEnCombo cMonedaT, CLng(paMonedaPesos)
    Foco cMonedaT
End Sub

Private Sub bPreview_Click()
    AccionImprimir
End Sub

Private Sub cConvertir_Click()
    
    If cConvertir.Value = vbChecked Then
        cMonedaN.Enabled = True: cMonedaN.BackColor = Obligatorio
        BuscoCodigoEnCombo cMonedaN, CLng(paMonedaPesos)
    Else
        cMonedaN.Enabled = False: cMonedaN.BackColor = Inactivo
        cMonedaN.Text = ""
    End If
    
End Sub

Private Sub cConvertir_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then If cMonedaN.Enabled Then Foco cMonedaN Else bConsultar.SetFocus
End Sub

Private Sub cMonedaN_GotFocus()
    cMonedaN.SelStart = 0
    cMonedaN.SelLength = Len(cMonedaN.Text)
End Sub

Private Sub cMonedaN_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cMonedaN.ListIndex > -1 Then bConsultar.SetFocus
End Sub

Private Sub Form_Activate()
    Me.Refresh
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    FechaDelServidor
    IncicializoControles
    
    cConvertir.Value = vbUnchecked
    cMonedaN.Enabled = False: cMonedaN.BackColor = Inactivo
    
    Screen.MousePointer = 0
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Error al inicializar el formulario.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function ValidoDatos() As Boolean

    ValidoDatos = False
    
    If cMonedaT.ListIndex = -1 Then
        MsgBox "Debe seleccionar la moneda para totalizar las operaciones.", vbExclamation, "ATENCIÓN"
        Foco cMonedaT: Exit Function
    End If
    
    If cConvertir.Value = vbChecked And cMonedaN.ListIndex = -1 Then
        MsgBox "Debe seleccionar la moneda nacional para buscar la tasa de cambio.", vbExclamation, "ATENCIÓN"
        Foco cMonedaN: Exit Function
    End If
    
    ValidoDatos = True
    
End Function

Private Sub Form_Resize()
On Error Resume Next

    If Me.WindowState = vbMinimized Then Exit Sub
    With Frame1
        .Top = 20
        .Left = 60
        .Width = Me.ScaleWidth - (.Left * 2)
    End With
    
    With picBotones
        .Top = Me.ScaleHeight - .Height
        .Left = 0
        .Width = Me.ScaleWidth
        .BorderStyle = 0
    End With
    
    With vsLista
        .Left = Frame1.Left
        .Width = Frame1.Width
        .Top = Frame1.Top + Frame1.Height + 60
        .Height = Me.ScaleHeight - (.Top + picBotones.Height)
    End With
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EndMain
End Sub

Private Sub Label1_Click()
    Foco cMonedaT
End Sub

Private Sub MnuCambiarBase_Click()
    Dim newB As String
    On Error GoTo errCh
    If MsgBox("Ud. sabe lo que está haciendo !!?", vbQuestion + vbYesNo + vbDefaultButton2, "Realmente desea cambiar la base") = vbNo Then Exit Sub
    
    newB = InputBox("Ingrese el texto del login para la nueva conexión" & vbCrLf & _
                "Id de aplicación en archivo de conexiones.", "Cambio de Base de Datos")
    
    If Trim(newB) = "" Then Exit Sub
    If MsgBox("Está seguro de cambiar la base de datos al login " & newB, vbQuestion + vbYesNo + vbDefaultButton2, "Cambiar Base") = vbNo Then Exit Sub
    
    newB = miConexion.TextoConexion(newB)
    If Trim(newB) = "" Then Exit Sub
    
    Screen.MousePointer = 11
    On Error Resume Next
    cBase.Close
    On Error GoTo errCh
    Set cBase = Nothing
    Set cBase = eBase.OpenConnection("", rdDriverNoPrompt, , newB)
    
    Screen.MousePointer = 0
    
    MsgBox "Ahora está trabajanbo en la nueva base de datos.", vbExclamation, "Base Cambiada OK"
    Exit Sub
    
errCh:
    clsGeneral.OcurrioError "Error de Conexión", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuOpConsultar_Click()
    AccionConsultar
End Sub

Private Sub MnuOpImprimir_Click()
    AccionImprimir
End Sub

Private Sub MnuSalFormulario_Click()
    Unload Me
End Sub

Private Sub cMonedaT_GotFocus()
    cMonedaT.SelStart = 0: cMonedaT.SelLength = Len(cMonedaT.Text)
End Sub

Private Sub cMonedaT_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then cConvertir.SetFocus

End Sub

Private Sub AccionImprimir(Optional ToPrinter As Boolean = False)

    On Error GoTo errPrint
    With orPrev
        .Caption = "Totalizador de Operaciones"
        .Header = "Totalizador de Operaciones"
        .FileName = "Totalizador de Operaciones"
        .AddGrid vsLista.hwnd
        
        If ToPrinter Then
            .GoPrint
        Else
            .ShowPreview
        End If
        
    End With
    
    Exit Sub

errPrint:
    clsGeneral.OcurrioError "Error al imprirmir. " & Trim(Err.Description)
    Screen.MousePointer = 0

End Sub

Private Function ArmoConsulta(TipoDeCredito As Integer, Moneda As Long, Optional CGSA As Boolean = True) As String

    Cons = "Select Count(*) as Cantidad, Sum(CreSaldoFactura) as Importe " & _
                " From Credito, Documento" & _
                " Where CreTipo = " & TipoDeCredito & _
                " And CreSaldoFactura > 0 " & _
                " And CreFactura = DocCodigo" & _
                " And DocMoneda = " & Moneda & _
                " And DocAnulado = 0"
                
    If CGSA Then Cons = Cons & " And CreMega = 0" Else Cons = Cons & " And CreMega = 1"
    
    ArmoConsulta = Cons
    
End Function

Private Sub IncicializoControles()
On Error Resume Next

    Cons = "Select MonCodigo, MonSigno From Moneda Order by MonSigno"
    CargoCombo Cons, cMonedaT, ""
    CargoCombo Cons, cMonedaN, ""
    BuscoCodigoEnCombo cMonedaT, CLng(paMonedaPesos)

    With vsLista
        .Editable = False
        .Rows = 1: .Cols = 4
        .FormatString = "<Tipo de Operaciones|> Cantidad|>Saldo|>Saldo M/N|"
        .ColWidth(0) = 3000: .ColWidth(1) = 1000: .ColWidth(2) = 1750:: .ColWidth(3) = 1750
        
        .AllowUserResizing = flexResizeColumns
        .ExtendLastCol = True
        .AllowBigSelection = False
        .AllowSelection = False
        
        .FocusRect = flexFocusNone
        .HighLight = flexHighlightNever
        .ScrollBars = flexScrollBarVertical
        .MergeCells = flexMergeSpill
        
        .AutoSizeMode = flexAutoSizeRowHeight
        .RowHidden(0) = True
        .BackColorBkg = .BackColor
    End With
   
    With img1
        bConsultar.Picture = .ListImages("consultar").ExtractIcon
        bImprimir.Picture = .ListImages("print").ExtractIcon
        bNoFiltros.Picture = .ListImages("limpiar").ExtractIcon
        bCancelar.Picture = .ListImages("salir").ExtractIcon
        bPreview.Picture = .ListImages("vista1").ExtractIcon
    End With
       
End Sub

Private Function zAddTitulo(mDeQuien As String)

    With vsLista
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = "Operaciones Totalizadas al: " & Format(Date, "Ddd d Mmm yyyy") & " - " & mDeQuien
        .Cell(flexcpFontBold, .Rows - 1, 0, , .Cols - 1) = True
        .Cell(flexcpFontSize, .Rows - 1, 0, , .Cols - 1) = 10
        .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 2) = Colores.Gris
        
        .AddItem " "
        
        'Veo si hay que sacar la TC a la moneda nacional
        If cMonedaN.Enabled Then
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = "Tasa de Cambio " & cMonedaT.Text & " -> " & cMonedaN.Text & " al " & Format(Date, "d Mmm yyyy") & ": " & Format(aTC, "#,##0.000")
            .Cell(flexcpFontItalic, .Rows - 1, 0) = True
            .AddItem ""
        End If
        
    End With
    
End Function

Private Function zAddTotal(mDeQuien As String)

    With vsLista
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = "Total General " & mDeQuien
        .Cell(flexcpText, .Rows - 1, 1) = Format(aCantidadT, "#,##0")
        .Cell(flexcpText, .Rows - 1, 2) = Format(aImporteT, "#,##0.00")
    
        If cMonedaN.Enabled Then .Cell(flexcpText, .Rows - 1, 3) = Format(aImporteT * aTC, "#,##0.00")
        
        .Cell(flexcpFontBold, .Rows - 1, 0, , .Cols - 1) = True
        .Cell(flexcpBackColor, .Rows - 1, 1, , 2) = Colores.Azul
        .Cell(flexcpForeColor, .Rows - 1, 1, , 2) = Colores.Blanco
    End With
    
End Function
