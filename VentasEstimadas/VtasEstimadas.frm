VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Begin VB.Form VtasEstimadas 
   Caption         =   "Ventas Estimadas"
   ClientHeight    =   5940
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VtasEstimadas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8LCtl.VSFlexGrid vsReales 
      Height          =   1455
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Width           =   6975
      _cx             =   12303
      _cy             =   2566
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
      GridLinesFixed  =   4
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   315
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
   Begin VSFlex8LCtl.VSFlexGrid gVtasEstimadas 
      Height          =   1455
      Left            =   120
      TabIndex        =   19
      Top             =   1500
      Width           =   6975
      _cx             =   12303
      _cy             =   2566
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
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   315
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
   Begin VB.Timer tmStart 
      Interval        =   30
      Left            =   6600
      Top             =   600
   End
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
      _Version        =   196608
      _ExtentX        =   1931
      _ExtentY        =   873
      _StockProps     =   229
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
   End
   Begin VB.Frame fBotones 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Width           =   6975
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1440
         MaxLength       =   12
         TabIndex        =   6
         Top             =   0
         Width           =   1215
      End
      Begin VB.TextBox tPorcentaje 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3840
         MaxLength       =   12
         TabIndex        =   8
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton bCancelar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   6600
         Picture         =   "VtasEstimadas.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Salir. [Ctrl+X]"
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bLimpiar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   6240
         Picture         =   "VtasEstimadas.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpiar el formulario. [Ctrl+L]"
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bCopiar 
         Height          =   310
         Left            =   5520
         Picture         =   "VtasEstimadas.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Copiar Ventas. [Ctrl+E]"
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   5880
         Picture         =   "VtasEstimadas.frx":0948
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Imprimir. [Ctrl+I]"
         Top             =   0
         Width           =   310
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha Tope Real:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Porcentaje:"
         Height          =   255
         Left            =   2880
         TabIndex        =   7
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.OptionButton opOpcion 
      Appearance      =   0  'Flat
      Caption         =   "&Artículo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton opOpcion 
      Appearance      =   0  'Flat
      Caption         =   "&Grupo"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox tArticulo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   4095
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   5685
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4471
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin AACombo99.AACombo cGrupo 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   240
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   556
      Style           =   2
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
   Begin VB.Image imgGrupo 
      Height          =   240
      Left            =   0
      Picture         =   "VtasEstimadas.frx":0A4A
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgArticulo 
      Height          =   240
      Left            =   0
      Picture         =   "VtasEstimadas.frx":0E47
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgGrupo1 
      Height          =   240
      Left            =   0
      Picture         =   "VtasEstimadas.frx":1234
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgGrupo2 
      Height          =   240
      Left            =   6720
      Picture         =   "VtasEstimadas.frx":12DD
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Copiar Ventas"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   6975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Reales"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   6975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Estimadas"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   6975
   End
   Begin VB.Label labArticulos 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   6615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      Height          =   975
      Left            =   120
      Top             =   120
      Width           =   6975
   End
   Begin VB.Menu MnuVentas 
      Caption         =   "GrillaVentas"
      Visible         =   0   'False
      Begin VB.Menu MnuAddMemo 
         Caption         =   "Insertar comentario"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuAñoSuperior 
         Caption         =   "&Insertar Año Siguiente"
      End
   End
   Begin VB.Menu MnuGruposMemos 
      Caption         =   "Grupos"
      Visible         =   0   'False
      Begin VB.Menu MnuGruposItem 
         Caption         =   "-"
         Index           =   0
      End
   End
End
Attribute VB_Name = "VtasEstimadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim avisos As VBControlExtender
Dim WithEvents avisosE As ComentariosAplicacion.ucAviso
Attribute avisosE.VB_VarHelpID = -1

'RDO.----------------------------
Private Rs As rdoResultset
Private aTituloTabla As String, aFormato As String

Private Sub InvocoComentarioEstimadas()
    If CStr(gVtasEstimadas.Cell(flexcpText, gVtasEstimadas.RowSel, gVtasEstimadas.ColSel)) = "" Or gVtasEstimadas.RowSel <= 0 Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Dim iCol As Integer
    Dim iRow As Integer
    iRow = gVtasEstimadas.RowSel
    iCol = gVtasEstimadas.ColSel
    avisos.Left = gVtasEstimadas.Cell(flexcpLeft, iRow, iCol)
    avisos.Top = gVtasEstimadas.Top + gVtasEstimadas.Cell(flexcpTop, iRow - 1, iCol) '+ gVtasEstimadas.Cell(flexcpHeight, iRow, iCol)
    If (opOpcion(0).Value) Then
        CambioAPP 73, tArticulo.Tag, Format(CDate("01/" & gVtasEstimadas.ColSel & "/" & gVtasEstimadas.Cell(flexcpValue, gVtasEstimadas.RowSel, 0)), "MM/yyyy"), "", 9, True
    Else
        CambioAPP 74, cGrupo.ItemData(cGrupo.ListIndex), Format(CDate("01/" & gVtasEstimadas.ColSel & "/" & gVtasEstimadas.Cell(flexcpValue, gVtasEstimadas.RowSel, 0)), "MM/yyyy"), cGrupo.Text, 11, True
    End If
    Screen.MousePointer = vbDefault
    
End Sub
Private Sub CambioAPP(ByVal app As Integer, ByVal idApp As Long, ByVal Mes As String, ByVal grupo As String, ByVal temaDefault As Integer, ByVal editar As Boolean)
On Error Resume Next
    'public void InitializeControl2(int aplicacion, int idAplicacion, int usuario, string idApp2)
    avisos.object.InitializeControl2 app, idApp, miConexion.UsuarioLogueado(True), Mes
    avisos.object.TituloConIDApp = grupo & IIf(grupo <> "", " - ", "") & "Mes: " & Mes
    avisos.object.temaDefault = temaDefault
    avisos.object.DisableIngreso = Not editar
    avisos.object.AbrirFormularioAplicacion
End Sub

Private Sub CrearAvisos()
On Error GoTo errCA
    If (avisos.Tag <> "") Then Exit Sub
    avisos.Tag = "1"
    avisos.object.InitializeControl2 73, 0, miConexion.UsuarioLogueado(True), ""
    avisos.object.TituloConIDApp = "Ventas estimadas"
    avisos.object.ShowGlobal = False
    avisos.object.Titulo = "Mes:"
    avisos.object.OcultarCfgTema = True
    avisos.object.EstiloComentarios = 1
    avisos.Visible = True
    avisos.Enabled = False
errCA:
End Sub

Sub InicializoControlAvisos()
On Error GoTo errStartAvisos
    Set avisos = Controls.Add("ComentariosAplicacion.ucAviso", "avisos", Me)
    Set avisos.Container = Me
    avisos.Left = 0
    avisos.Width = 200 ' Me.Width - 300
    avisos.Top = 0
    avisos.Height = 0 'picComentarios.Height
    avisos.Visible = False
    avisos.Tag = ""
    Set avisosE = avisos.object
    Exit Sub
errStartAvisos:
End Sub

Private Sub avisosE_DataChanged()
    
    If opOpcion(0).Value Then
        BuscoVentas tArticulo.Tag, True
    Else
        BuscoVentas cGrupo.Tag, False
    End If
End Sub

Private Sub bCancelar_Click()
On Error Resume Next
    Unload Me
End Sub

Private Sub bCopiar_Click()
On Error GoTo ErrBC
    'Verifico que ingresó un artículo o un grupo de artículos.
    If opOpcion(0).Value = True Then
        If tArticulo.Tag = "" Then
            MsgBox "No se ingresó un artículo.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
    Else
        If cGrupo.ListIndex = -1 Then
            MsgBox "No se ingresó un grupo de artículos.", vbExclamation, "ATENCIÓN"
            cGrupo.Tag = ""
            Exit Sub
        Else
            If cGrupo.Tag = "" Then
                MsgBox "El grupo no posee artículos asociados, esta consulta no retorna información.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
        End If
    End If
    
    If Not IsNumeric(tPorcentaje.Text) Then MsgBox "No se ingreso un valor numérico en el porcentaje.", vbExclamation, "ATENCIÓN": tPorcentaje.SetFocus: Exit Sub
    If Not IsDate(tFecha.Text) Then MsgBox "No se ingreso una fecha válida.", vbExclamation, "ATENCIÓN": tFecha.SetFocus: Exit Sub
    If Val(tPorcentaje.Text) < -100 Then MsgBox "Se registrará un valor negativo si ingresa cantidades negativas mayores que cien.", vbInformation, "ATENCIÓN": Exit Sub
    
    'Hay ventas reales.---------------
    If vsReales.Rows = 1 Then MsgBox "No hay ventas reales a copiar.", vbInformation, "ATENCIÓN"
    'Válido Fechas.--------------------
    If PrimerDia(CDate(tFecha.Text)) >= PrimerDia(Format(gFechaServidor, "d/mm/yyyy")) Then MsgBox "La fecha tope debe ser inferior al primer día del mes corriente.", vbInformation, "ATENCIÓN": Exit Sub
    If CDate(tFecha.Text) < PrimerDia(gFechaServidor - 365) Then
        If MsgBox("La fecha tope ingresada copiara ventas que no son consideradas en calculos del stock." & Chr(13) & "¿Desea continuar de todas formas?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
    End If
    RelojA
    If opOpcion(0).Value = True Then CopioVentas tArticulo.Tag Else CopioVentas cGrupo.Tag
    If opOpcion(0).Value = True Then LimpioCampos: LimpioGrilla: BuscoVentas tArticulo.Tag, True Else LimpioCampos: LimpioGrilla: BuscoVentas cGrupo.Tag, False
    RelojD
    Exit Sub
ErrBC:
    MsgBox "Error :" & Err.Description, vbCritical, "Error"
    
End Sub
Private Sub CopioVentas(ByVal Articulos As String)
On Error GoTo ErrCV
    Do While Articulos <> ""
        If InStr(Articulos, ",") > 0 Then
            CopioVentasArticulo (Mid(Articulos, 1, InStr(Articulos, ",") - 1))
            Articulos = Mid(Articulos, InStr(Articulos, ",") + 1, Len(Articulos))
        Else
            CopioVentasArticulo Articulos
            Articulos = ""
        End If
    Loop
    Exit Sub
ErrCV:
    MensajeError "Ocurrió un error al invocar la copia de ventas.", Trim(Err.Description)
    RelojD
End Sub
Private Sub CopioVentasArticulo(Articulo As String)
On Error GoTo ErrCV

    'Busco las ventas reales.-----------------------
    Cons = "Select Mes = DatePart(mm,AArFecha), Ano = DatePart(yy,AArFecha), Cantidad = (Sum(AArCantidadNCo) + Sum(AArCantidadNCr) + Sum(AArCantidadECo) + Sum(AArCantidadECr))" _
        & " From AcumuladoArticulo " _
        & " Where AArArticulo = " & Articulo _
        & "And AArFEcha Between '" & Format(Format(PrimerDia(tFecha.Text), "dd/mm") & "/" & Year(tFecha.Text) - 1, sqlFormatoF) & "'" _
        & " And '" & Format(UltimoDia(tFecha.Text), sqlFormatoF) & "'" _
        & " Group by DatePart(mm,AArFecha), DatePart(yy,AArFecha)"
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly)
    Do While Not Rs.EOF
        If Rs!Cantidad > 0 Then
            Cons = "Select * From VentasEstimadas Where VEsArticulo = " & Articulo _
                & " And VEsMesAño = '" & Format("01/" & Rs!Mes & "/" & Rs!Ano + 1, sqlFormatoF) & "'"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If RsAux.EOF Then
                If Rs!Cantidad + CInt(((Rs!Cantidad * tPorcentaje.Text) / 100)) > 0 Then
                    RsAux.AddNew
                    RsAux!VEsarticulo = Articulo
                    RsAux!VesMesAño = Format("01/" & Rs!Mes & "/" & Rs!Ano + 1, sqlFormatoF)
                    RsAux!VEsCantidad = Rs!Cantidad + CInt(((Rs!Cantidad * tPorcentaje.Text) / 100))
                    RsAux.Update
                End If
            Else
                If Rs!Cantidad + CInt(((Rs!Cantidad * tPorcentaje.Text) / 100)) > 0 Then
                    RsAux.Edit
                    RsAux!VEsCantidad = Rs!Cantidad + CInt(((Rs!Cantidad * tPorcentaje.Text) / 100))
                    RsAux.Update
                Else
                    RsAux.Delete
                End If
            End If
            RsAux.Close
        End If
        Rs.MoveNext
    Loop
    Rs.Close
    Exit Sub
ErrCV:
    MensajeError "Ocurrió un error al copiar las ventas reales.", Err.Description
    RelojD
End Sub

Private Sub bImprimir_Click()
Dim J As Integer, aTexto As String

    If gVtasEstimadas.Rows = 1 Then
        MsgBox "No hay datos en la lista para realizar la impresión.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    On Error GoTo errPrint
    RelojA
    aTituloTabla = ""
    
    With vsListado
    
        If Not .PrintDialog(pdPrinterSetup) Then Screen.MousePointer = 0: Exit Sub
        
        .Preview = True
        .StartDoc
                
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión.", vbInformation, "ATENCIÓN"
            Screen.MousePointer = vbDefault: Exit Sub
        End If
    
        EncabezadoListado vsListado, "Importaciones - Consulta de Ventas Estimadas Vs. Ventas Reales", False
        
        .FileName = "Consulta de Ventas Estimadas"
        .FontSize = 8: .FontBold = False
        
        For I = 1 To gVtasEstimadas.Rows - 1
            aTexto = Trim(gVtasEstimadas.Cell(flexcpText, I, 0))
            For J = 1 To 12
                aTexto = aTexto & "|" & Trim(gVtasEstimadas.Cell(flexcpText, I, J))
            Next
            .AddTable aFormato, "", aTexto, Colores.Inactivo, , True
        Next
        
        .Paragraph = ""
        .Paragraph = ""
        
        With vsListado
            .FontSize = 8
            .FontBold = True
            .TableBorder = tbBoxRows
            .AddTable aFormato, aTituloTabla, "", , Colores.Inactivo
            .FontBold = False
        End With

        For I = 1 To vsReales.Rows - 1
            aTexto = Trim(vsReales.Cell(flexcpText, I, 0))
            For J = 1 To 12
                aTexto = aTexto & "|" & Trim(vsReales.Cell(flexcpText, I, J))
            Next
            .AddTable aFormato, "", aTexto, Colores.Inactivo, , True
        Next
        
        .EndDoc
        .PrintDoc True
        
    End With
    
    RelojD
    Exit Sub

errPrint:
    RelojD
    MensajeError "Ocurrió un error al realizar la impresión. ", Err.Description

End Sub

Private Sub bLimpiar_Click()
On Error Resume Next
    RelojA
    LimpioCampos
    LimpioGrilla
    labArticulos.Caption = ""
    tArticulo.Text = ""
    cGrupo.ListIndex = -1
    If opOpcion(0).Value Then Foco tArticulo Else Foco cGrupo
    RelojD
End Sub
Private Sub cGrupo_Click()
On Error GoTo ErrGC
    LimpioCampos
    LimpioGrilla
    labArticulos.Caption = ""
    'Detallo los artículos que pertenecen al grupo
    If cGrupo.ListIndex > -1 Then
        RelojA
        Cons = "Select Artid, ArtNombre From ArticuloGrupo, Articulo Where AGrGrupo = " & cGrupo.ItemData(cGrupo.ListIndex) _
            & " And ArtEnUso <> 0  And AGrArticulo = ArtID"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
        If RsAux.EOF Then
            labArticulos.Caption = "No hay artículos en ese grupo."
            cGrupo.Tag = ""
        Else
            Do While Not RsAux.EOF
                If labArticulos.Caption = "" Then
                    cGrupo.Tag = RsAux!ArtID
                    labArticulos.Caption = Trim(RsAux!ArtNombre)
                Else
                    cGrupo.Tag = cGrupo.Tag & ", " & RsAux!ArtID
                    labArticulos.Caption = labArticulos.Caption & ", " & Trim(RsAux!ArtNombre)
                End If
                RsAux.MoveNext
            Loop
        End If
        RsAux.Close
        RelojD
    End If
    Exit Sub
ErrGC:
    RelojD
    MensajeError "Ocurrió un error al buscar los artículos del grupo.", Err.Description
End Sub
Private Sub cGrupo_GotFocus()
On Error GoTo ErrGr
    With cGrupo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda " Seleccione el grupo de artículos a consultar."
    If cGrupo.ListCount = 0 Then
        Cons = "Select GruCodigo, GruNombre From Grupo Order by GruNombre"
        CargoCombo Cons, cGrupo, ""
    End If
    Exit Sub
ErrGr:
    MensajeError "Ocurrió un error inesperado.", Trim(Err.Description)
End Sub
Private Sub cGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cGrupo.ListIndex > -1 And Trim(cGrupo.Tag) <> "" Then LimpioGrilla: BuscoVentas cGrupo.Tag, False
End Sub
Private Sub cGrupo_LostFocus()
    Ayuda ""
    cGrupo.SelStart = 0
End Sub
Private Sub Form_Activate()
    RelojD
    Me.Refresh
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyX
                Unload Me
            Case vbKeyE
                bCopiar_Click
            Case vbKeyI
                bImprimir_Click
            Case vbKeyL
                bLimpiar_Click
        End Select
    Else
        Select Case KeyCode
            Case vbKeyEscape
                Unload Me
        End Select
    End If
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad
Dim Co As String

    RelojA
    ObtengoSeteoForm Me, HeightIni:=8000, WidthIni:=9000
    
    CargoParametrosImportaciones
    
    FechaDelServidor
    
    'Inicializo Controles de Artículos y Grupos.----------
    tArticulo.Left = cGrupo.Left: tArticulo.Height = cGrupo.Height: tArticulo.Width = cGrupo.Width
    LimpioCampos
    LimpioGrilla
    If Trim(Command()) <> "" Then
        If Mid(Command(), 1, 1) = "A" Then
            opOpcion(0).Value = True
            Cons = "Select * From Articulo Where ArtID =" & Mid(Command(), 2, Len(Command()))
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
            If Not RsAux.EOF Then
                tArticulo.Text = Trim(RsAux!ArtNombre)
                tArticulo.Tag = RsAux!ArtID
                RsAux.Close
                BuscoVentas (tArticulo.Tag), opOpcion(0).Value
            Else
                RsAux.Close
            End If
        Else
            'Viene un grupo.
            opOpcion(1).Value = True
            Cons = "Select GruCodigo, GruNombre From Grupo Order by GruNombre"
            CargoCombo Cons, cGrupo, ""
            BuscoCodigoEnCombo cGrupo, Mid(Command(), 2, Len(Command()))
            cGrupo_Click
            cGrupo_KeyPress (vbKeyReturn)
        End If
    End If
    InicializoControlAvisos
    Exit Sub
    
ErrLoad:
    MensajeError "Ocurrió un error al iniciar el formulario.", Trim(Err.Description)
    RelojD
    Exit Sub
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Ayuda ""
End Sub
Private Sub Form_Resize()
On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    fBotones.Left = 120
    fBotones.Width = Me.ScaleWidth - 240
    Label1.Width = Me.ScaleWidth - 220
    
    
    Label2.Width = Label1.Width: Label3.Width = Label1.Width
    Shape1.Width = Label1.Width
    labArticulos.Width = Label1.Width - 100
    gVtasEstimadas.Width = Label1.Width
    vsReales.Width = Label1.Width
    Label1.Top = Shape1.Top + Shape1.Height + 60
    gVtasEstimadas.Top = Label1.Top + Label1.Height
    gVtasEstimadas.Height = (Me.ScaleHeight - (Label1.Top + Label1.Height + Label2.Height + Label3.Height + Status.Height + fBotones.Height + Shape1.Height)) / 1.5
    Label3.Top = gVtasEstimadas.Top + gVtasEstimadas.Height + 80
    fBotones.Top = Label3.Top + Label3.Height + 40
    Label2.Top = fBotones.Top + fBotones.Height + 40
    vsReales.Top = Label2.Top + Label2.Height + 40
    vsReales.Height = Me.ScaleHeight - (Shape1.Top + Label1.Height + Label2.Height + Label3.Height + Status.Height + fBotones.Height + Shape1.Height + gVtasEstimadas.Height + 310)
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    CierroConexion
    GuardoSeteoForm Me
    Set clsGeneral = Nothing
End Sub

Private Sub gVtasEstimadas_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
    If gVtasEstimadas.Cell(flexcpForeColor, Row, col) <> RGB(140, 0, 0) Then Cancel = True
    If opOpcion(1).Value Then Cancel = True
End Sub

Private Sub gVtasEstimadas_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo errBM

    If Button = vbRightButton Then Exit Sub

    ' get cell that was clicked
    Dim r&, C&
    r = gVtasEstimadas.MouseRow
    C = gVtasEstimadas.MouseCol
    If r <= 0 Then Exit Sub
    gVtasEstimadas.Select r, C
    
    ' make sure the click was on the sheet
    If CStr(gVtasEstimadas.Cell(flexcpText, r, C)) = "" Or r < 0 Or C < 0 Then Exit Sub
    
    ' make sure the click was on a cell with a button
    If Not (gVtasEstimadas.Cell(flexcpPicture, r, C) Is IIf(opOpcion(0).Value, imgGrupo, imgArticulo)) Then Exit Sub
      
    ' make sure the click was on the button (not just on the cell)
    ' note: this works for right-aligned buttons
    Dim d!
    d = gVtasEstimadas.Cell(flexcpLeft, r, C) + imgGrupo.Width
    If (X > d) Then Exit Sub
    
    avisos.Left = gVtasEstimadas.Cell(flexcpLeft, r, C)
    avisos.Top = gVtasEstimadas.Top + gVtasEstimadas.Cell(flexcpTop, r - 1, C)
    
    Dim idGrupo As Integer
    Dim nomGrupo As String
    If (opOpcion(0).Value) Then
        idGrupo = BuscoGrupoDeArticuloConComentario(Format(CDate("01/" & C & "/" & gVtasEstimadas.Cell(flexcpValue, r, 0)), "MM/yyyy"), nomGrupo)
    Else
        idGrupo = BuscoArticulosEnGrupoConComentario(Format(CDate("01/" & C & "/" & gVtasEstimadas.Cell(flexcpValue, r, 0)), "MM/yyyy"), nomGrupo)
    End If
    If idGrupo > 0 Then
        CambioAPP IIf(opOpcion(0).Value, 74, 73), idGrupo, Format(CDate("01/" & C & "/" & gVtasEstimadas.Cell(flexcpValue, r, 0)), "MM/yyyy"), nomGrupo, IIf(opOpcion(0).Value, 9, 11), False
        'InvocoComentarioDeGrupoEnArticulo idGrupo, Format(CDate("01/" & C & "/" & gVtasEstimadas.Cell(flexcpValue, r, 0)), "MM/yyyy"), nomGrupo, True
    End If
    ' cancel default processing
    ' note: this is not strictly necessary in this case, because
    '       the dialog box already stole the focus etc, but let's be safe.
    Cancel = True
errBM:
End Sub

Function BuscoArticulosEnGrupoConComentario(ByVal Mes As String, ByRef nomArticulo As String) As Long
On Error GoTo errGB
    
    BuscoArticulosEnGrupoConComentario = 0
    Do While MnuGruposItem.Count > 1
        Unload MnuGruposItem(MnuGruposItem.UBound)
    Loop
    
    Dim sInsertados As String
    Dim iQ As Integer
    iQ = 0
    
    Cons = "SELECT ArtID, ArtNombre FROM TemasAplicacionesComentarios INNER JOIN ArticuloGrupo ON TACIDEntidad = AGrArticulo AND AGrGrupo = " & cGrupo.ItemData(cGrupo.ListIndex) _
        & " INNER JOIN Articulo ON ArtID = AGrArticulo " _
        & " WHERE TACTema IN (9, 10) AND TACIDEntidad2 = '" & Mes & "' ORDER BY ArtNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        If (InStr(1, "," & sInsertados & ",", "," & RsAux("ArtID") & ",") = 0) Then
            sInsertados = IIf(sInsertados <> "", ",", "") & RsAux("ArtID")
            If (iQ > 0) Then
                Load MnuGruposItem(MnuGruposItem.UBound + 1)
                BuscoArticulosEnGrupoConComentario = 0
            End If
            iQ = iQ + 1
            MnuGruposItem(MnuGruposItem.UBound).Caption = Trim(RsAux("ArtNombre"))
            MnuGruposItem(MnuGruposItem.UBound).Tag = RsAux("ArtID") & "|" & Mes
            If (iQ = 1) Then
                BuscoArticulosEnGrupoConComentario = RsAux("ArtID")
                nomArticulo = Trim(RsAux("ArtNombre"))
            End If
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    If (MnuGruposItem.Count > 1) Then
        PopupMenu MnuGruposMemos
    End If
    Exit Function
errGB:
    BuscoArticulosEnGrupoConComentario = -1
End Function

Function BuscoGrupoDeArticuloConComentario(ByVal Mes As String, ByRef nomGrupo As String) As Long
On Error GoTo errGB
    
    BuscoGrupoDeArticuloConComentario = 0
    Do While MnuGruposItem.Count > 1
        Unload MnuGruposItem(MnuGruposItem.UBound)
    Loop
    
    Dim sInsertados As String
    Dim iQ As Integer
    iQ = 0
    
    Cons = "SELECT GruCodigo, GruNombre FROM TemasAplicacionesComentarios INNER JOIN ArticuloGrupo ON TACIDEntidad = AGrGrupo " _
        & "AND AGrArticulo = " & tArticulo.Tag & "  INNER JOIN Grupo ON AGrGrupo = GruCodigo WHERE TACTema IN (11, 12) AND TACIDEntidad2 = '" & Mes & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        If (InStr(1, "," & sInsertados & ",", "," & RsAux("GruCodigo") & ",") = 0) Then
            sInsertados = IIf(sInsertados <> "", ",", "") & RsAux("GruCodigo")
            If (iQ > 0) Then
                Load MnuGruposItem(MnuGruposItem.UBound + 1)
                BuscoGrupoDeArticuloConComentario = 0
            End If
            iQ = iQ + 1
            MnuGruposItem(MnuGruposItem.UBound).Caption = Trim(RsAux("GruNombre"))
            MnuGruposItem(MnuGruposItem.UBound).Tag = RsAux("GruCodigo") & "|" & Mes
            If (iQ = 1) Then
                BuscoGrupoDeArticuloConComentario = RsAux("GruCodigo")
                nomGrupo = Trim(RsAux("GruNombre"))
            End If
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    If (MnuGruposItem.Count > 1) Then
        PopupMenu MnuGruposMemos
    End If
    Exit Function
errGB:
    BuscoGrupoDeArticuloConComentario = -1
End Function

'Sub InvocoComentarioDeGrupoEnArticulo(ByVal idGrupo As Long, ByVal Mes As String, ByVal grupo As String, ByVal Estimada As Boolean)
'
'    Screen.MousePointer = vbHourglass
'    'Busco los grupos que tienen comentario.
'    CambioAPP 74, idGrupo, Mes, grupo, IIf(Estimada, 11, 12)
'    Screen.MousePointer = vbDefault
'
'End Sub


Private Sub gVtasEstimadas_DblClick()
    InvocoComentarioEstimadas
End Sub

Private Sub gVtasEstimadas_LostFocus()
    gVtasEstimadas.Select 0, 0, 0, 0
End Sub

Private Sub gVtasEstimadas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gVtasEstimadas.Rows > 1 And Button = vbRightButton Then
        PopupMenu MnuVentas
        Exit Sub
    End If
End Sub

Private Sub gVtasEstimadas_RowColChange()
    On Error Resume Next
    
'    If gVtasEstimadas.Cell(flexcpData, gVtasEstimadas.RowSel, gVtasEstimadas.ColSel) Is Nothing Then
'        lblComentario.Caption = gVtasEstimadas.Cell(flexcpData, gVtasEstimadas.RowSel, gVtasEstimadas.ColSel)
'        lblComentario.Visible = (lblComentario.Caption <> "")
'    End If
End Sub

Private Sub gVtasEstimadas_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
On Error GoTo ErrGV
    
    If Not IsNumeric(gVtasEstimadas.EditText) Then Cancel = True: Exit Sub
    If opOpcion(1).Value Then Exit Sub
    RelojA
    If IsNumeric(gVtasEstimadas.EditText) Then
        Cons = "Select * From VentasEstimadas Where VEsMesAño = '" & Format("01/" & col & "/" & gVtasEstimadas.Cell(flexcpValue, Row, 0), sqlFormatoF) & "'" _
            & " And VEsArticulo = " & tArticulo.Tag
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux.EOF Then
            If Val(gVtasEstimadas.EditText) > 0 Then
                RsAux.AddNew
                RsAux!VEsarticulo = tArticulo.Tag
                RsAux!VesMesAño = Format("01/" & col & "/" & gVtasEstimadas.Cell(flexcpValue, Row, 0), sqlFormatoF)
                RsAux!VEsCantidad = gVtasEstimadas.EditText
                RsAux.Update
            Else
                gVtasEstimadas.EditText = ""
            End If
        Else
            If Val(gVtasEstimadas.EditText) > 0 Then
                RsAux.Edit
                RsAux!VEsCantidad = gVtasEstimadas.EditText
                RsAux.Update
            Else
                RsAux.Delete
                gVtasEstimadas.EditText = ""
            End If
        End If
        RsAux.Close
    End If
    RelojD
    Exit Sub
ErrGV:
    MensajeError "Ocurrió un error al intentar modificar las ventas estimadas.", Err.Description
    RelojD
End Sub
Private Sub Label4_Click()
    Foco tFecha
End Sub
Private Sub Label5_Click()
    Foco tPorcentaje
End Sub

Private Sub MnuAddMemo_Click()
On Error GoTo errAM
    Dim memo As String
    memo = InputBox("Ingrese el comentario.", "Comentario")
    If Trim(memo) <> "" Then
        Cons = "SELECT * FROM VentasEstimadas WHERE VEs"
    End If
    Exit Sub
errAM:
End Sub

Private Sub MnuAñoSuperior_Click()
On Error GoTo ErrMAS
    RelojA
    gVtasEstimadas.AddItem "", 1
    gVtasEstimadas.Cell(flexcpText, 1, 0) = gVtasEstimadas.Cell(flexcpText, 2, 0) + 1
    gVtasEstimadas.Cell(flexcpForeColor, 1, 1, 1, 12) = RGB(140, 0, 0)
    RelojD
    Exit Sub
ErrMAS:
    MensajeError "Ocurrió un error al intentar insertar una fila en la grilla.", Trim(Err.Description)
    RelojD
End Sub

Private Sub MnuGruposItem_Click(Index As Integer)
Dim vInfo() As String
    vInfo = Split(MnuGruposItem(Index).Tag, "|")
    'puse por defecto estimada pero hay que controlar la grilla
    'IIf(opOpcion(0).Value, 9, 11)
    CambioAPP IIf(opOpcion(0).Value, 74, 73), vInfo(0), vInfo(1), MnuGruposItem(Index).Caption, 0, False
    'InvocoComentarioDeGrupoEnArticulo vInfo(0), vInfo(1), MnuGruposItem(Index).Caption, True
End Sub

Private Sub opOpcion_Click(Index As Integer)
    labArticulos.Caption = ""
    LimpioCampos
    LimpioGrilla
    If Index = 0 Then
        cGrupo.Visible = False
        tArticulo.Visible = True: tArticulo.Text = "": tArticulo.Tag = ""
    Else
        cGrupo.Visible = True: cGrupo.Text = "": cGrupo.Tag = ""
        tArticulo.Visible = False
    End If
End Sub
Private Sub opOpcion_GotFocus(Index As Integer)
    Ayuda " Seleccione el tipo de consulta."
End Sub
Private Sub opOpcion_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 0 Then tArticulo.SetFocus Else cGrupo.SetFocus
    End If
End Sub
Private Sub opOpcion_LostFocus(Index As Integer)
    Ayuda ""
End Sub
Private Sub Ayuda(strMensaje As String)
    Status.Panels(4).Text = strMensaje
End Sub
Private Sub tArticulo_Change()
    tArticulo.Tag = ""
    LimpioCampos
    LimpioGrilla
End Sub
Private Sub tArticulo_GotFocus()
    With tArticulo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda " Ingrese el artículo a consultar."
End Sub
Private Sub tArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(tArticulo.Text) <> "" Then
        LimpioGrilla
        If IsNumeric(tArticulo.Text) Then
            BuscoArticuloPorCodigo CLng(tArticulo.Text)
        Else
            BuscoArticuloPorNombre
        End If
        If Trim(tArticulo.Tag) <> "" Then BuscoVentas (tArticulo.Tag), True
    End If
End Sub
Private Sub tArticulo_LostFocus()
    tArticulo.SelStart = 0
    Ayuda ""
End Sub
Private Sub BuscoArticuloPorCodigo(Articulo As Long)
On Error GoTo ErrBAPC
    RelojA
    Cons = "Select ArtID, ArtNombre, ArtEnUso From Articulo Where ArtCodigo = " & CLng(Articulo)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If RsAux.EOF Then
        RelojD
        MsgBox "No existe un artículo con ese código, o el mismo fue eliminado.", vbInformation, "ATENCIÓN"
        LimpioCampos
    Else
        If RsAux!ArtEnUso <> 0 Then
            tArticulo.Text = Trim(RsAux!ArtNombre): tArticulo.Tag = RsAux!ArtID
        Else
            MsgBox "El artículo seleccionado no esta habilitado.", vbInformation, "ATENCIÓN"
        End If
    End If
    RsAux.Close
    RelojD
    Exit Sub
ErrBAPC:
    MensajeError "Ocurrió un error al buscar el artículo por código.", Err.Description
    RelojD
End Sub
Private Sub BuscoArticuloPorNombre()
    Cons = "Select ArtCodigo, 'Código' = ArtCodigo, Nombre = ArtNombre" _
        & " From Articulo" _
        & " Where ArtNombre LIKE '" & Replace(Trim(tArticulo.Text), " ", "%") & "%'" _
        & " And ArtEnUso <> 0"
    PresentoListaDeAyuda Cons
End Sub
Private Sub PresentoListaDeAyuda(strConsulta As String)
On Error GoTo ErrPLDA
Dim Resultado As String
    RelojA
    
    'Limpio los valores del textbox.
    tArticulo.Tag = "": tArticulo.Text = ""
    
    Dim sqlAyuda As New clsListadeAyuda
    If sqlAyuda.ActivarAyuda(cBase, Cons, 4500, 1) > 0 Then
    'Obtengo si hay seleccionado.---------------
        Resultado = sqlAyuda.RetornoDatoSeleccionado(0)
    End If
    'Destruyo la clase.------------------------------
    Set sqlAyuda = Nothing
    RelojA
    If Resultado <> "" Then
        If IsNumeric(Resultado) Then
           BuscoArticuloPorCodigo CLng(Resultado)
        Else
            RelojD
            MsgBox "Se espera que se retorne el código de artículo.", vbInformation, "ATENCIÓN"
        End If
    End If
    RelojD
    Exit Sub
ErrPLDA:
    RelojD
    MensajeError "Ocurrió un error al presentar la lista de ayuda.", Err.Description
End Sub
Private Sub LimpioCampos()
    tFecha.Text = ""
    tPorcentaje.Text = ""
End Sub
Private Sub LimpioGrilla()
On Error GoTo ErrLG
Dim iCol As Integer
    'Grilla.-------------------------------
    With gVtasEstimadas
        .Redraw = False
        .ExtendLastCol = False
        .Clear
        .Editable = True
        .Rows = 1
        '.FormatString = "^Año/Mes|Enero |Febrero |Marzo |Abril |Mayo |Junio |Julio |Agosto |Setiembre |Octubre |Noviembre |Diciembre |>Total"
        .FormatString = "^Año/Mes|Enero    |Febrero    |Marzo    |Abril    |Mayo    |Junio   |Julio   |Agosto   |Setiembre  |Octubre  |Noviembre  |Diciembre  |>Total"
        For iCol = 1 To 13
            .ColWidth(iCol) = 900
        Next
        .AllowUserResizing = flexResizeColumns
        .AllowBigSelection = False
        .AllowSelection = False
        .SelectionMode = flexSelectionFree
        .Select 0, 0, 0, 0
        .Redraw = True
        .RowHeightMin = 315
    End With
    With vsReales
        .Redraw = False
        .ExtendLastCol = False
        .Clear
        .Editable = False
        .AllowBigSelection = False
        .AllowSelection = False
        .Rows = 1
        .RowHeightMin = 315
        .FormatString = "^Año/Mes|Enero    |Febrero    |Marzo    |Abril    |Mayo    |Junio   |Julio   |Agosto   |Setiembre  |Octubre  |Noviembre  |Diciembre  |>Total"
        For iCol = 1 To 13
            .ColWidth(iCol) = 900
        Next
        .AllowUserResizing = flexResizeColumns
        .Select 0, 0, 0, 0
        .SelectionMode = flexSelectionFree
        .Redraw = True
    End With
    '--------------------------------------
    
    Exit Sub
ErrLG:
    MensajeError "Ocurrió un error al inicializar las grillas.", Trim(Err.Description)
    RelojD
End Sub
Private Sub BuscoVentas(ByVal Articulos As String, ByVal unoSolo As Boolean)
On Error GoTo ErrBVE
Dim Año As String, fila As Integer
    
    RelojA
    Año = "": fila = 0
    LimpioCampos
    
    'Busco las ventas estimadas.-----------------
    If unoSolo Then
    '& ", (SELECT COUNT(DISTINCT(AGrGrupo)) FROM TemasAplicacionesComentarios INNER JOIN ArticuloGrupo ON TACIDEntidad = AGrGrupo AND AGrArticulo = VEsArticulo WHERE TACTema IN (12) AND TACIDEntidad2 = Right('0'+ RTrim(Convert(char(2), Month(VesMesAño))), 2) + '/' + CONVERT(Char(4), Year(VesMesAño))) as VentasGR "
        Cons = "Select VesMesAño, Cantidad = VEsCantidad " _
            & ", (SELECT COUNT(*) FROM TemasAplicacionesComentarios WHERE TACIDEntidad = VEsarticulo AND TACTema IN (9) AND TACIDEntidad2 = Right('0'+ RTrim(Convert(char(2), Month(VesMesAño))), 2) + '/' + CONVERT(Char(4), Year(VesMesAño))) as VentasE " _
            & ", (SELECT COUNT(*) FROM TemasAplicacionesComentarios WHERE TACIDEntidad = VEsarticulo AND TACTema IN (10) AND TACIDEntidad2 = Right('0'+ RTrim(Convert(char(2), Month(VesMesAño))), 2) + '/' + CONVERT(Char(4), Year(VesMesAño))) as VentasR " _
            & ", (SELECT COUNT(DISTINCT(AGrGrupo)) FROM TemasAplicacionesComentarios INNER JOIN ArticuloGrupo ON TACIDEntidad = AGrGrupo AND AGrArticulo = VEsArticulo WHERE TACTema IN (11, 12) AND TACIDEntidad2 = Right('0'+ RTrim(Convert(char(2), Month(VesMesAño))), 2) + '/' + CONVERT(Char(4), Year(VesMesAño))) as VentasG" _
            & " FROM VentasEstimadas " _
            & " Where VEsArticulo = " & Articulos _
            & " Order by VEsMesAño DESC"
    Else
        Cons = "Select VesMesAño, Cantidad = Sum(VEsCantidad) " _
            & ", (SELECT COUNT(*) FROM TemasAplicacionesComentarios WHERE TACIDEntidad = " & cGrupo.ItemData(cGrupo.ListIndex) & " AND TACTema IN (11) AND TACIDEntidad2 = Right('0'+ RTrim(Convert(char(2), Month(VesMesAño))), 2) + '/' + CONVERT(Char(4), Year(VesMesAño))) as VentasE " _
            & ", (SELECT COUNT(*) FROM TemasAplicacionesComentarios WHERE TACIDEntidad = " & cGrupo.ItemData(cGrupo.ListIndex) & " AND TACTema IN (12) AND TACIDEntidad2 = Right('0'+ RTrim(Convert(char(2), Month(VesMesAño))), 2) + '/' + CONVERT(Char(4), Year(VesMesAño))) as VentasR " _
            & ", (SELECT COUNT(DISTINCT(AGrGrupo)) FROM TemasAplicacionesComentarios INNER JOIN ArticuloGrupo ON TACIDEntidad = AGrArticulo AND AGrGrupo = " & cGrupo.ItemData(cGrupo.ListIndex) & " WHERE TACTema IN (9, 10) AND TACIDEntidad2 = Right('0'+ RTrim(Convert(char(2), Month(VesMesAño))), 2) + '/' + CONVERT(Char(4), Year(VesMesAño))) as VentasG" _
            & " From VentasEstimadas Where VEsArticulo IN (" & Articulos & ")" _
            & " Group by VEsMesAño" _
            & " Order by VEsMesAño DESC"
    End If
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly)
    
    If Not RsAux.EOF Then
        If Val(Format(RsAux!VesMesAño, "yyyy")) < Val(Format(gFechaServidor, "yyyy")) + 1 Then
            Año = Format(gFechaServidor, "yyyy") + 2
        Else
            Año = Format(RsAux!VesMesAño, "yyyy")
        End If
        Do While Not RsAux.EOF
            If Val(Año) <> Format(RsAux!VesMesAño, "yyyy") Then
                'Inserto la fila.---------------------------------------------
                gVtasEstimadas.AddItem ""
                Año = Val(Año) - 1
                fila = fila + 1
                gVtasEstimadas.TextMatrix(fila, 0) = Val(Año)
                
                If Val(Año) > Year(gFechaServidor) Then
                    gVtasEstimadas.Cell(flexcpForeColor, fila, 1, fila, 12) = RGB(140, 0, 0)
                ElseIf Val(Año) = Year(gFechaServidor) Then
                    gVtasEstimadas.Cell(flexcpForeColor, fila, Month(gFechaServidor), fila, 12) = RGB(140, 0, 0)
                End If
                'Inserte el año mayor a la que tengo en el resulset.
            Else
                If fila = 0 Then
                    gVtasEstimadas.AddItem ""
                    fila = fila + 1
                    gVtasEstimadas.TextMatrix(fila, 0) = Val(Año)
                    If Year(RsAux!VesMesAño) > Year(gFechaServidor) Then
                        gVtasEstimadas.Cell(flexcpForeColor, fila, 1, fila, 12) = RGB(140, 0, 0)
                    ElseIf Year(RsAux!VesMesAño) = Year(gFechaServidor) Then
                        gVtasEstimadas.Cell(flexcpForeColor, fila, Month(gFechaServidor), fila, 12) = RGB(140, 0, 0)
                    End If
                End If
                gVtasEstimadas.TextMatrix(fila, Month(RsAux!VesMesAño)) = RsAux!Cantidad
                If RsAux("VentasE") > 0 Then
                    gVtasEstimadas.Cell(flexcpBackColor, fila, Month(RsAux("VesMesAño"))) = &HDDFFFF
                ElseIf RsAux("VentasR") > 0 Then
                    gVtasEstimadas.Cell(flexcpBackColor, fila, Month(RsAux("VesMesAño"))) = &HCCE8FF
                End If
                
'                If (unoSolo) Then
                    'En grupos no existe esta col x eso no es un AND.
                    If (RsAux("VentasG") > 0) Then
                        gVtasEstimadas.Cell(flexcpPicture, fila, Month(RsAux("VesMesAño"))) = IIf(opOpcion(0).Value, imgGrupo, imgArticulo)
                        gVtasEstimadas.Cell(flexcpPictureAlignment, fila, Month(RsAux("VesMesAño"))) = flexAlignLeftCenter
                    End If
'                End If
                RsAux.MoveNext
            End If
        Loop
        RsAux.Close
        
        Dim cTotal As Currency
        cTotal = 0
        Dim col As Byte
        'Recorro las cols de cada fila para presentar el total.
        For fila = 1 To gVtasEstimadas.Rows - 1
            cTotal = 0
            For col = 1 To gVtasEstimadas.Cols - 2
                cTotal = cTotal + Val(gVtasEstimadas.Cell(flexcpText, fila, col))
            Next
            If cTotal <> 0 Then gVtasEstimadas.Cell(flexcpText, fila, gVtasEstimadas.Cols - 1) = cTotal
        Next
        
    Else
        RsAux.Close
        'No hay ventas estimadas.------------------
        gVtasEstimadas.AddItem ""
        fila = fila + 1
        gVtasEstimadas.Cell(flexcpForeColor, fila, 1, fila, 12) = RGB(140, 0, 0)
        gVtasEstimadas.TextMatrix(fila, 0) = Format(gFechaServidor, "yyyy") + 1
        gVtasEstimadas.AddItem ""
        fila = fila + 1
        gVtasEstimadas.TextMatrix(fila, 0) = Format(gFechaServidor, "yyyy")
        gVtasEstimadas.Cell(flexcpForeColor, fila, Month(gFechaServidor), fila, 12) = RGB(140, 0, 0)
        Año = Year(gFechaServidor)
    End If
    
    If CInt(Año) > Year(gFechaServidor) Then
        gVtasEstimadas.AddItem ""
        fila = gVtasEstimadas.Rows - 1
        gVtasEstimadas.TextMatrix(fila, 0) = Format(gFechaServidor, "yyyy")
        gVtasEstimadas.Cell(flexcpForeColor, fila, Month(gFechaServidor), fila, 12) = RGB(140, 0, 0)
    End If
    
    
    Año = "": fila = 0
    
    If unoSolo Then
        '& ", (SELECT COUNT(DISTINCT(AGrGrupo)) FROM TemasAplicacionesComentarios INNER JOIN ArticuloGrupo ON TACIDEntidad = AGrGrupo AND AGrArticulo = VEsArticulo WHERE TACTema IN (11, 12) AND TACIDEntidad2 = Right('0'+ RTrim(Convert(char(2), Month(VesMesAño))), 2) + '/' + CONVERT(Char(4), Year(VesMesAño))) as VentasG"
        Cons = "Select Mes = DatePart(mm,AArFecha), Ano = DatePart(yy,AArFecha), Cantidad = (Sum(AArCantidadNCo) + Sum(AArCantidadNCr) + Sum(AArCantidadECo) + Sum(AArCantidadECr))" _
            & ", (SELECT COUNT(*) FROM TemasAplicacionesComentarios WHERE TACIDEntidad = AArArticulo AND TACTema IN (9) AND TACIDEntidad2 = Right('0'+ RTrim(Convert(char(2), Month(AArFecha))), 2) + '/' + CONVERT(Char(4), Year(AArFecha))) as VentasE " _
            & ", (SELECT COUNT(*) FROM TemasAplicacionesComentarios WHERE TACIDEntidad = AArArticulo AND TACTema IN (10) AND TACIDEntidad2 = Right('0'+ RTrim(Convert(char(2), Month(AArFecha))), 2) + '/' + CONVERT(Char(4), Year(AArFecha))) as VentasR " _
            & ", (SELECT COUNT(DISTINCT(AGrGrupo)) FROM TemasAplicacionesComentarios INNER JOIN ArticuloGrupo ON TACIDEntidad = AGrGrupo AND AGrArticulo = AArArticulo WHERE TACTema IN (11, 12) AND TACIDEntidad2 = Right('0'+ RTrim(Convert(char(2), Month(AArFecha))), 2) + '/' + CONVERT(Char(4), Year(AArFecha))) as VentasG" _
            & " From AcumuladoArticulo " _
            & " Where AArArticulo = " & Articulos _
            & "And AArFEcha <'" & Format(PrimerDia(gFechaServidor), sqlFormatoF) & "'" _
            & " Group by AArArticulo, DatePart(mm,AArFecha), DatePart(yy,AArFecha) Order By Ano Desc, Mes Desc"
    Else
        'Busco las ventas reales.-----------------------
        
        Cons = "Select Mes = DatePart(mm,AArFecha), Ano = DatePart(yy,AArFecha), Cantidad = (Sum(AArCantidadNCo) + Sum(AArCantidadNCr) + Sum(AArCantidadECo) + Sum(AArCantidadECr))" _
            & ", (SELECT COUNT(*) FROM TemasAplicacionesComentarios WHERE TACIDEntidad = " & cGrupo.ItemData(cGrupo.ListIndex) & " AND TACTema IN (11) AND TACIDEntidad2 = Right('0'+ RTrim(Convert(char(2), Month(AArFecha))), 2) + '/' + CONVERT(Char(4), Year(AArFecha))) as VentasE " _
            & ", (SELECT COUNT(*) FROM TemasAplicacionesComentarios WHERE TACIDEntidad = " & cGrupo.ItemData(cGrupo.ListIndex) & " AND TACTema IN (12) AND TACIDEntidad2 = Right('0'+ RTrim(Convert(char(2), Month(AArFecha))), 2) + '/' + CONVERT(Char(4), Year(AArFecha))) as VentasR " _
            & ", (SELECT COUNT(DISTINCT(AGrArticulo)) FROM TemasAplicacionesComentarios INNER JOIN ArticuloGrupo ON TACIDEntidad = AGrArticulo AND AGrGrupo = " & cGrupo.ItemData(cGrupo.ListIndex) & " WHERE TACTema IN (9, 10) AND TACIDEntidad2 = Right('0'+ RTrim(Convert(char(2), Month(AArFecha))), 2) + '/' + CONVERT(Char(4), Year(AArFecha))) as VentasG" _
            & " From AcumuladoArticulo " _
            & " Where AArArticulo IN (" & Articulos & ")" _
            & "And AArFEcha <'" & Format(PrimerDia(gFechaServidor), sqlFormatoF) & "'" _
            & " Group by DatePart(mm,AArFecha), DatePart(yy,AArFecha) Order By Ano Desc, Mes Desc"
    End If
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly)
    
    If Not RsAux.EOF Then
        Do While Not RsAux.EOF
            If Val(Año) <> RsAux!Ano Then
                Año = RsAux!Ano
                vsReales.AddItem ""
                fila = fila + 1
                vsReales.TextMatrix(fila, 0) = RsAux!Ano
            End If

'            'Veo si tiene comentario.-----------
'            Cons = "Select CVRComentario From ComentarioVentaReal Where CVRArticulo IN (" & Articulos & ")" _
'                & " And CVRFecha = '" & Format("01/" & RsAux!mes & "/" & RsAux!Ano, sqlFormatoF) & "'"
'            Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
'            If Not Rs.EOF Then
'                vsReales.Cell(flexcpForeColor, Fila, RsAux!mes) = vbBlue
'            End If
'            Rs.Close
            
            If RsAux("VentasR") > 0 Then
                vsReales.Cell(flexcpBackColor, fila, RsAux("mes")) = &HCCE8FF
            ElseIf RsAux("VentasE") > 0 Then
                vsReales.Cell(flexcpBackColor, fila, RsAux("mes")) = &HDDFFFF
            End If
'            If (unoSolo) Then
                'En grupos no existe esta col x eso no es un AND.
                If (RsAux("VentasG") > 0) Then
                    vsReales.Cell(flexcpPicture, fila, RsAux("mes")) = IIf(opOpcion(0).Value, imgGrupo, imgArticulo)
                    vsReales.Cell(flexcpPictureAlignment, fila, RsAux("mes")) = flexAlignLeftCenter
                End If
'            End If

            vsReales.TextMatrix(fila, RsAux!Mes) = RsAux!Cantidad
            RsAux.MoveNext
        Loop
        
        For fila = 1 To vsReales.Rows - 1
            cTotal = 0
            For col = 1 To vsReales.Cols - 2
                cTotal = cTotal + Val(vsReales.Cell(flexcpText, fila, col))
            Next
            If cTotal <> 0 Then vsReales.Cell(flexcpText, fila, vsReales.Cols - 1) = cTotal
        Next
        
    End If
    RsAux.Close
    '-------------------------------------------------------
    RelojD
    Exit Sub
ErrBVE:
    MensajeError "Ocurrió un error al buscar las ventas estimadas.", Trim(Err.Description)
    RelojD
End Sub

Private Sub tFecha_GotFocus()
    With tFecha
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese la última fecha real a considerar."
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsDate(tFecha.Text) Then tPorcentaje.SetFocus
    End If
End Sub

Private Sub tFecha_LostFocus()
    If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, FormatoFP)
    Ayuda ""
End Sub

Private Sub tmStart_Timer()
    tmStart.Enabled = False
    CrearAvisos
End Sub

Private Sub tPorcentaje_GotFocus()
    With tPorcentaje
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese un porcentaje a incrementar o disminuir al copiar las ventas estimadas."
End Sub

Private Sub tPorcentaje_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tPorcentaje.Text) Then bCopiar.SetFocus
    End If
End Sub

Private Sub tPorcentaje_LostFocus()
    Ayuda ""
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub vsListado_NewPage()
    vsListado.Paragraph = ""
    vsListado.Paragraph = ""
    If opOpcion(0).Value Then
        vsListado.Paragraph = "Artiículo: " & tArticulo.Text
    Else
        vsListado.Paragraph = "Artiículos: " & labArticulos.Caption
    End If
    vsListado.Paragraph = ""
    
    If aTituloTabla = "" Then ImpresionEncabezadoTabla
    With vsListado
        .FontSize = 8
        .FontBold = True
        .TableBorder = tbBoxRows
        .AddTable aFormato, aTituloTabla, "", , Colores.Inactivo
        .FontBold = False
    End With

End Sub

Private Sub vsReales_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    If Button = vbRightButton Then Exit Sub

    ' get cell that was clicked
    Dim r&, C&
    r = vsReales.MouseRow
    C = vsReales.MouseCol
    vsReales.Select r, C
    
    ' make sure the click was on the sheet
    If CStr(vsReales.Cell(flexcpText, r, C)) = "" Or r < 0 Or C < 0 Then Exit Sub
    
    ' make sure the click was on a cell with a button
    If Not (vsReales.Cell(flexcpPicture, r, C) Is IIf(opOpcion(0).Value, imgGrupo, imgArticulo)) Then Exit Sub
      
    ' make sure the click was on the button (not just on the cell)
    ' note: this works for right-aligned buttons
    Dim d!
    d = vsReales.Cell(flexcpLeft, r, C) + imgGrupo.Width
    If (X > d) Then Exit Sub
    
    avisos.Left = vsReales.Cell(flexcpLeft, r, C)
    avisos.Top = vsReales.Top + vsReales.Cell(flexcpTop, r - 1, C)
    
    Dim idGrupo As Integer
    Dim nomGrupo As String
    
    
    If (opOpcion(0).Value) Then
        idGrupo = BuscoGrupoDeArticuloConComentario(Format(CDate("01/" & C & "/" & vsReales.Cell(flexcpValue, r, 0)), "MM/yyyy"), nomGrupo)
    Else
        idGrupo = BuscoArticulosEnGrupoConComentario(Format(CDate("01/" & C & "/" & vsReales.Cell(flexcpValue, r, 0)), "MM/yyyy"), nomGrupo)
    End If
    If idGrupo > 0 Then
        'InvocoComentarioDeGrupoEnArticulo idGrupo, Format(CDate("01/" & C & "/" & vsReales.Cell(flexcpValue, r, 0)), "MM/yyyy"), nomGrupo, False
        CambioAPP IIf(opOpcion(0).Value, 74, 73), idGrupo, Format(CDate("01/" & C & "/" & vsReales.Cell(flexcpValue, r, 0)), "MM/yyyy"), nomGrupo, IIf(opOpcion(0).Value, 10, 12), False
    End If
    Cancel = True
End Sub

Private Sub vsReales_DblClick()
On Error GoTo ErrCo
    'Veo si la celda seleccionada es azul hay datos.
'    If vsReales.Rows > 1 Then
'        If opOpcion(0).Value Then
'            RelojA
'            ComVtaReal.pSeleccionado = tArticulo.Tag
'            ComVtaReal.pMes = vsReales.col & "/" & vsReales.Cell(flexcpText, vsReales.Row, 0)
'            ComVtaReal.Show vbModeless, Me
'            RelojD
'        Else
'            RelojA
'            Cons = "Select ArtCodigo,  Artículo = '(' + RTRim(CONVERT(Char(10), ArtCodigo)) + ')' + ' ' + ArtNombre, Comentario = CVRComentario" _
'                & " From Articulo, ComentarioVentaReal" _
'                & " Where ArtID IN(" & cGrupo.Tag & ")" _
'                & " And CVRFecha = '" & Format("01/" & vsReales.col & "/" & vsReales.Cell(flexcpText, vsReales.Row, 0), sqlFormatoF) & "'" _
'                & " And ArtID = CVRArticulo"
'            Dim frmAyuda As New clsListadeAyuda
'            frmAyuda.ActivarAyuda cBase, Cons, 7000, 0, "Ayuda"
'            Set frmAyuda = Nothing
'            RelojD
'        End If
'    End If

    If CStr(vsReales.Cell(flexcpText, vsReales.RowSel, vsReales.ColSel)) = "" Or vsReales.RowSel <= 0 Or vsReales.RowSel >= vsReales.Cols Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Dim iCol As Integer
    Dim iRow As Integer
    iRow = vsReales.RowSel
    iCol = vsReales.ColSel
    avisos.Left = vsReales.Cell(flexcpLeft, iRow, iCol)
    avisos.Top = vsReales.Top + vsReales.Cell(flexcpTop, iRow - 1, iCol)
    If (opOpcion(0).Value) Then
        CambioAPP 73, tArticulo.Tag, Format(CDate("01/" & vsReales.ColSel & "/" & vsReales.Cell(flexcpValue, iRow, 0)), "MM/yyyy"), "", 10, True
    Else
        CambioAPP 74, cGrupo.ItemData(cGrupo.ListIndex), Format(CDate("01/" & vsReales.ColSel & "/" & vsReales.Cell(flexcpValue, vsReales.RowSel, 0)), "MM/yyyy"), cGrupo.Text, 12, True
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrCo:
    MensajeError "Ocurrió un error al presentar el comentario.", Err.Description
    RelojD
End Sub

Private Sub ImpresionEncabezadoTabla()
    
    'El encabezado es el mismo para ambas grillas.----------------------------------------
    aTituloTabla = "": aFormato = ""
    For I = 0 To vsReales.Cols - 1
        vsReales.Row = 0
        Select Case vsReales.ColAlignment(0)
            Case lvwColumnCenter: aFormato = aFormato & "+^~"
            Case lvwColumnLeft: aFormato = aFormato & "+<~"
            Case lvwColumnRight: aFormato = aFormato & "+>~"
        End Select
        aFormato = aFormato & CInt(vsReales.ColWidth(I) * 1.5) & "|"
        aTituloTabla = aTituloTabla & vsReales.TextMatrix(0, I) & "|"
    Next
    aFormato = Mid(aFormato, 1, Len(aFormato) - 1)
    aTituloTabla = Mid(aTituloTabla, 1, Len(aTituloTabla) - 1)
    
End Sub

Private Sub vsReales_LostFocus()
    vsReales.Select 0, 0, 0, 0
End Sub
