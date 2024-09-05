VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form frmPreciosFlete 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Precio de Flete"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7980
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBotton 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   7980
      TabIndex        =   4
      Top             =   3855
      Width           =   7980
      Begin VB.CommandButton butCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   6480
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton butOK 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4920
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblPrecioSel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Precio seleccionado:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   30
         Width           =   1335
      End
      Begin VB.Label lblPreSugerido 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Precio sugerido:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label lblInfoPrecioSugerido 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio sugerido:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label lblInfoPrecioSeleccionado 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio seleccionado:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   30
         Width           =   1815
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid lstDatos 
      Align           =   1  'Align Top
      Height          =   1695
      Left            =   0
      TabIndex        =   2
      Top             =   1455
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   2990
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
   Begin VB.PictureBox picFiltros 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   0
      ScaleHeight     =   960
      ScaleWidth      =   7980
      TabIndex        =   1
      Top             =   495
      Width           =   7980
      Begin VB.Label lblOrigen 
         BackStyle       =   0  'Transparent
         Caption         =   "Origen:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label lblFecha 
         BackStyle       =   0  'Transparent
         Caption         =   "Embarque:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   11
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label lblContenedores 
         BackStyle       =   0  'Transparent
         Caption         =   "Contenedores:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label lblLinea 
         BackStyle       =   0  'Transparent
         Caption         =   "Línea:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label lblAgencia 
         BackStyle       =   0  'Transparent
         Caption         =   "Agencia:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   8
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.PictureBox picTitulo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   7980
      TabIndex        =   0
      Top             =   0
      Width           =   7980
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio de Flete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmPreciosFlete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum colsGrilla
    CG_SiNo
    CG_APartir
    CG_Hasta
    CG_Contenedor
    CG_Cantidad
    CG_Unitario
    CG_Total
End Enum

Public DatosEmbarque As clsDatosPrecioFlete
Public DialogResult As Boolean
Public PrecioSeleccionado As Currency

Private Sub CargoGrilla()
    lstDatos.Rows = 1
    Cons = "SELECT FEmFAPartir, FEmFHasta, FEmImporte, FEmContenedor FROM FleteEmbarque " _
        & " WHERE FEmOrigen = " & DatosEmbarque.Origen.Codigo _
        & " AND FEmDestino = " & DatosEmbarque.Destino.Codigo _
        & " And FEmAgencia = " & DatosEmbarque.Agencia.Codigo _
        & " And FEmContenedor IN (" & DatosEmbarque.ContenedoresIDs & ")" _
        & " And (FEmFAPartir <= '" & Format(DatosEmbarque.FechaEmbarque, "yyyymmdd 00:00:00") & "' OR FEMFAPartir Is Null) " _
        & " AND FEmFHasta + 1 >= '" & Format(DatosEmbarque.FechaEmbarque, "yyyymmdd 23:59:00") & "'" _
        & " And FEmLinea = " & DatosEmbarque.Linea.Codigo _
        & " ORDER BY FEmContenedor, FEMFAPartir DESC"

    Dim idContAnt As Long
    Dim impSugerido As Currency
    
    Dim rsFE As rdoResultset
    Set rsFE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsFE.EOF
        With lstDatos
            .AddItem ""
            If Not IsNull(rsFE("FEmFAPartir")) Then .Cell(flexcpText, .Rows - 1, CG_APartir) = Format(rsFE("FEmFAPartir"), "dd/MM/yy")
            If Not IsNull(rsFE("FEmFHasta")) Then .Cell(flexcpText, .Rows - 1, CG_Hasta) = Format(rsFE("FEmFHasta"), "dd/MM/yy")
            Dim oCEm As clsContenedoresEmbarque
            Set oCEm = DatosEmbarque.Contenedor(rsFE("FEmContenedor"))
            .Cell(flexcpText, .Rows - 1, CG_Contenedor) = oCEm.Contenedor.Nombre
            .Cell(flexcpData, .Rows - 1, CG_Contenedor) = oCEm.Contenedor.Codigo
            .Cell(flexcpText, .Rows - 1, CG_Cantidad) = oCEm.Cantidad
            .Cell(flexcpText, .Rows - 1, CG_Unitario) = Format(rsFE("FEmImporte"), "#,##0.00")
            .Cell(flexcpText, .Rows - 1, CG_Total) = Format(rsFE("FEmImporte") * oCEm.Cantidad, "#,##0.00")
            
            If idContAnt <> rsFE("FEmContenedor") Then
                idContAnt = rsFE("FEmContenedor")
                .Cell(flexcpChecked, .Rows - 1, CG_SiNo) = flexChecked
                impSugerido = impSugerido + (rsFE("FEmImporte") * oCEm.Cantidad)
            End If
            
        End With
        rsFE.MoveNext
    Loop
    rsFE.Close
    
    'Por defecto selecciono el de mayor fecha a partir ya que es la última contización que me dieron.
    lblPreSugerido.Caption = Format(impSugerido, "#,##0.00")
    AsignoPrecioSeleccionado

End Sub

Private Sub AsignoPrecioSeleccionado()
Dim iX As Integer
Dim impSel As Currency

    With lstDatos
        For iX = 1 To lstDatos.Rows - 1
            If .Cell(flexcpChecked, iX, CG_SiNo) = flexChecked Then
                impSel = impSel + CCur(.Cell(flexcpText, iX, CG_Total))
            End If
        Next
    End With
    lblPrecioSel.Caption = Format(impSel, "#,##0.00")
End Sub

Private Sub butCancelar_Click()
    Unload Me
End Sub

Private Sub butOK_Click()
    If Not EstanTodosClickeados Then
        MsgBox "Algunos contenedores no fueron clickeados, verifique.", vbExclamation, "Posible error"
        Exit Sub
    End If
    If MsgBox("¿Confirma asignar al precio de flete el seleccionado?", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
        PrecioSeleccionado = CCur(lblPrecioSel.Caption)
        DialogResult = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    InicializoGrilla
    DialogResult = False
    If DatosEmbarque Is Nothing Then Exit Sub
    lblOrigen.Caption = "Origen: " & DatosEmbarque.Origen.Nombre
    lblAgencia.Caption = "Agencia: " & DatosEmbarque.Agencia.Nombre
    lblLinea.Caption = "Línea: " & DatosEmbarque.Linea.Nombre
    lblFecha.Caption = IIf(DatosEmbarque.SiEmbarco, "Embarcó: ", "Previsto: ") & DatosEmbarque.FechaEmbarque
    If Not DatosEmbarque.SiEmbarco Then lblFecha.ForeColor = ColorNaranja
    lblContenedores.Caption = "Contenedores: " & DatosEmbarque.ContenedoresALineaTexto()
    
    CargoGrilla
End Sub

Private Sub InicializoGrilla()
    
    With lstDatos
        .Redraw = False
        .ExtendLastCol = False
        .Clear
        .BackColorBkg = vbWindowBackground
        .Editable = True
        .Rows = 1
        .Cols = 1
        .FixedCols = 0
        .ForeColor = vbBlack
        .FormatString = "Selección|A Partir|Rige hasta|Contenedor|>Cantidad|>Unitario|>Total"
        .ColDataType(0) = flexDTBoolean
        .ColWidth(1) = 1000
        .ColWidth(3) = 1200
        .ColWidth(5) = 1500
        .ColWidth(6) = 1500
        .ColFormat(5) = "#,##0.00"
        .ColFormat(6) = "#,##0.00"
        .AllowUserResizing = flexResizeColumns
        .ExtendLastCol = True
        .Redraw = True
    End With

End Sub

Private Sub Form_Resize()
On Error Resume Next
    lstDatos.Height = Me.ScaleHeight - (lstDatos.Top + picBotton.Height)
    butCancelar.Left = Me.ScaleWidth - butCancelar.Width - 120
    butOK.Left = butCancelar.Left - butOK.Width - 120
End Sub

Private Function EstanTodosClickeados() As Boolean
Dim oCont As clsContenedoresEmbarque
Dim bClick As Boolean
    For Each oCont In DatosEmbarque.Contenedores
        'verifico que tenga en la grilla uno con clic.
        bClick = False
        Dim iX As Integer
        With lstDatos
            For iX = 1 To lstDatos.Rows - 1
                If .Cell(flexcpChecked, iX, CG_SiNo) = flexChecked And .Cell(flexcpData, iX, CG_Contenedor) = oCont.Contenedor.Codigo Then
                    bClick = True
                    Exit For
                End If
            Next
        End With
        If Not bClick Then
            EstanTodosClickeados = False
            Exit Function
        End If
    Next
    EstanTodosClickeados = True
End Function

Private Sub lstDatos_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    'Tomo el id del contenedor y deselecciono los otros del mismo contenedor.
    If (lstDatos.Cell(flexcpChecked, Row, Col) = flexChecked) Then
        Dim iX As Integer
        Dim impSel As Currency
        With lstDatos
            For iX = 1 To lstDatos.Rows - 1
                If iX <> Row And .Cell(flexcpChecked, iX, CG_SiNo) = flexChecked And .Cell(flexcpData, iX, CG_Contenedor) = .Cell(flexcpData, Row, CG_Contenedor) Then
                    .Cell(flexcpChecked, iX, CG_SiNo) = flexUnchecked
                End If
            Next
        End With
    End If
    AsignoPrecioSeleccionado
    
End Sub

Private Sub lstDatos_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Row < 1 Or Col > 0)
End Sub

