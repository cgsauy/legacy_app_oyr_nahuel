VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmSuceso 
   BackColor       =   &H8000000B&
   Caption         =   "Sucesos del Cliente"
   ClientHeight    =   4605
   ClientLeft      =   3285
   ClientTop       =   2205
   ClientWidth     =   7380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSuceso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   7380
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5741
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
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
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
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   " Cliente:"
      ForeColor       =   &H00FFFFFF&
      Height          =   250
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   675
   End
   Begin VB.Label lCliente 
      BackColor       =   &H00808080&
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   250
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   6255
   End
   Begin VB.Menu MnuBDerecho 
      Caption         =   "BotonDerecho"
      Visible         =   0   'False
      Begin VB.Menu MnuIrA 
         Caption         =   "Ir a ..."
      End
      Begin VB.Menu MnuVerL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSuceso 
         Caption         =   "Detalle del Suceso"
      End
      Begin VB.Menu MnuFactura 
         Caption         =   "Detalle de Factura"
      End
      Begin VB.Menu MnuComentarios 
         Caption         =   "Agregar Comentarios al Cliente"
      End
   End
End
Attribute VB_Name = "frmSuceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public prm_IdCliente As Long

Dim aValor As Long
Dim I As Integer

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    ObtengoSeteoForm Me, , , 10800
    
    Screen.MousePointer = 11
    InicializoGrilla
    
    If prm_IdCliente <> 0 Then
        BuscarCliente prm_IdCliente
        AccionConsultar
    End If
    
End Sub


Private Sub AccionConsultar()
On Error GoTo errPago
    
    Screen.MousePointer = 11
    
    'Query Con DATOS---------------------------------------------------------------------------------------------------------
    cons = "Select * from Suceso, Usuario" _
            & " Where SucCliente = " & prm_IdCliente _
            & " And SucUsuario = UsuCodigo"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    
    Do While Not rsAux.EOF
        With vsConsulta
            .AddItem CStr(rsAux!SucCodigo)
            .Cell(flexcpText, .Rows - 1, 1) = Format(rsAux!SucFecha, "dd/mm/yy hh:mm")
            
            If Not IsNull(rsAux!SucDocumento) Then aValor = rsAux!SucDocumento Else aValor = 0
            .Cell(flexcpData, .Rows - 1, 0) = aValor
            
            If Not IsNull(rsAux!SucCliente) Then aValor = rsAux!SucCliente Else aValor = 0
            .Cell(flexcpData, .Rows - 1, 1) = aValor

            
            
            If Not IsNull(rsAux!SucDescripcion) Then .Cell(flexcpText, .Rows - 1, 2) = Trim(rsAux!SucDescripcion)
            .Cell(flexcpText, .Rows - 1, 3) = Trim(rsAux!UsuIdentificacion)
                    
            If Not IsNull(rsAux!SucDefensa) Then .Cell(flexcpText, .Rows - 1, 4) = Trim(rsAux!SucDefensa)
            
        End With
        
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    Screen.MousePointer = 0
    If vsConsulta.Rows > 1 Then
        With vsConsulta
        .Select 1, 1: .Sort = flexSortGenericDescending
        
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 2, , False
        End With
    End If
    
    Exit Sub
    
errPago:
    clsGeneral.OcurrioError "Error al cargar los sucesos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    lCliente.Width = Me.ScaleWidth - lCliente.Left
    
    With vsConsulta
        .Width = Me.ScaleWidth: .Left = Me.ScaleLeft
        .Height = Me.ScaleHeight - .Top
    End With
    
    With vsConsulta
        Dim aSize As Currency
        For I = 0 To .Cols - 3: aSize = aSize + .ColWidth(I): Next I
        .ColWidth(.Cols - 2) = .Width - (aSize + .ColWidth(.Cols - 1) + 300)
        
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 2, , False
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    GuardoSeteoForm Me
    
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
End Sub

Private Sub MnuComentarios_Click()
On Error GoTo errCliente
    
    aValor = vsConsulta.Cell(flexcpData, vsConsulta.Row, 1)
    If aValor = 0 Then Exit Sub
    Screen.MousePointer = 11

    Dim miCliente As New clsCliente
    miCliente.Comentarios aValor
    
    Me.Refresh
    Set miCliente = Nothing
    Screen.MousePointer = 0
    Exit Sub
    
errCliente:
    clsGeneral.OcurrioError "Error al acceder a los comentarios.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuFactura_Click()
    On Error GoTo errFactura
        
    aValor = vsConsulta.Cell(flexcpData, vsConsulta.Row, 0)
    EjecutarApp App.Path & "\Detalle de Factura", CStr(aValor)
    
errFactura:
End Sub

Private Sub MnuSuceso_Click()
    Call vsConsulta_DblClick
End Sub


Private Sub InicializoGrilla()

    On Error Resume Next
    With vsConsulta
        .Cols = 1: .Rows = 1:
        .FormatString = "<Suceso|<Fecha|<Descripción|<Usuario|<Defensa|"
        .ColWidth(1) = 1260
        .ColWidth(2) = 3750: .ColWidth(3) = 1000: .ColWidth(4) = 3000
        
        .ColDataType(1) = flexDTDate
        .WordWrap = True: .ExtendLastCol = True
        
        .WordWrap = True
        .ColAlignment(0) = flexAlignLeftTop: .ColAlignment(1) = flexAlignLeftTop: .ColAlignment(2) = flexAlignLeftTop: .ColAlignment(3) = flexAlignLeftTop
        .ColAlignment(4) = flexAlignLeftTop
    End With
    
End Sub

Private Sub vsConsulta_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    On Error Resume Next
    
    Dim aSize As Currency
    With vsConsulta
        'For I = 0 To .Cols - 3: aSize = aSize + .ColWidth(I): Next I
        '.ColWidth(.Cols - 2) = .Width - (aSize + .ColWidth(.Cols - 1) + 200)
        
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 2, , False
    End With
    
End Sub


Private Sub vsConsulta_DblClick()
    If vsConsulta.Rows = 1 Then Exit Sub
    Screen.MousePointer = 11
    
    frmDetalle.prm_Suceso = vsConsulta.Cell(flexcpValue, vsConsulta.Row, 0)
    frmDetalle.Show vbModal, Me
    Me.Refresh
    
    Screen.MousePointer = 0
    
End Sub

Private Sub vsConsulta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo errBD
    With vsConsulta
        If .Rows = 1 Then Exit Sub
        If Button <> vbRightButton Then Exit Sub
        
        If .Cell(flexcpData, .Row, 0) = 0 Then MnuFactura.Enabled = False Else MnuFactura.Enabled = True
        If .Cell(flexcpData, .Row, 1) = 0 Then MnuComentarios.Enabled = False Else MnuComentarios.Enabled = True
        
        PopupMenu MnuBDerecho, , , , MnuIrA
    End With
    
errBD:
End Sub

Private Sub BuscarCliente(miId As Long)
    On Error GoTo errCCliente
    Dim aTexto As String
    
    cons = "Select Cliente.*, (RTrim(CPeNombre1) + ' ' + RTrim(isnull(CPeNombre2, '')) + ' ' + RTrim(CPeApellido1) + ' ' + RTrim(isnull(CPeApellido2, '')))  as Nombre" & _
               " From CPersona, Cliente Where CliCodigo = CPeCliente And CliCodigo = " & miId & _
                  " UNION ALL" & _
              " Select Cliente.*, (RTrim(CEmNombre) + ' (' + RTrim(isnull(CEmFantasia, '')) + ')')  as Nombre" & _
              " From CEmpresa, Cliente Where CliCodigo = CEmCliente And CliCodigo = " & miId
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        Select Case rsAux!CliTipo
            Case 1: If Not IsNull(rsAux!CliCiRuc) Then aTexto = clsGeneral.RetornoFormatoCedula(rsAux!CliCiRuc)
            Case 2: If Not IsNull(rsAux!CliCiRuc) Then aTexto = clsGeneral.RetornoFormatoRuc(Trim(rsAux!CliCiRuc))
        End Select
        
        lCliente.Caption = " " & aTexto & " " & Trim(rsAux!Nombre)
    Else
        lCliente.Caption = " No Existe !!"
    End If
    rsAux.Close
    Exit Sub

errCCliente:
End Sub
