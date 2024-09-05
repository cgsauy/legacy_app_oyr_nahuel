VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmAyuda 
   Caption         =   "Lista de Ayuda"
   ClientHeight    =   4815
   ClientLeft      =   3195
   ClientTop       =   1950
   ClientWidth     =   4365
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAyuda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4815
   ScaleWidth      =   4365
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   4140
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   8
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir de ayuda. [Ctrl+X]"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   350
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "seleccionar"
            Object.ToolTipText     =   "Seleccionar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "primero"
            Object.ToolTipText     =   "Primer página. [Ctrl+P]"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "anterior"
            Object.ToolTipText     =   "Página anterior. [Ctrl+A]"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "siguiente"
            Object.ToolTipText     =   "Siguiente página. [Ctrl+S]"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsAyuda 
      Height          =   2415
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   4260
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
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
      ExplorerBar     =   1
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
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4560
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList 
      Left            =   3480
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmAyuda.frx":0442
            Key             =   "primero"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmAyuda.frx":0794
            Key             =   "anterior"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmAyuda.frx":0AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmAyuda.frx":0DF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmAyuda.frx":118A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------
'   LISTA DE AYUDA (LiAyuda.frm)
'
'   Propiedades:
'
'       pSeleccionado: Retorna long con el valor de la KEY seleccionado. En el ejemplo CliCodigo
'       pSeleccionadoItem: Retorna el primer item de la lista. En el ejemplo CliCodigo (el segundo)
'
'                                                                                                                 A&A analistas 12-Nov 98
'                                                                                                                   Adaptación: 31-7-2000
'-----------------------------------------------------------------------------------------------------------------------
Option Explicit

Private PosLista As Integer, Cant As Integer
Const cCantidadMaxima = 100

Private intFilaSeleccionada As Integer
Public bolHidePrimera As Boolean

Property Get RetornoFilaSeleccionada() As Integer
    RetornoFilaSeleccionada = intFilaSeleccionada
End Property

Property Get RetornoDatoSeleccionado(intCol As Integer) As Variant
    'Retorna el dato que contiene la columna seleccionada.
    RetornoDatoSeleccionado = frmAyuda.vsAyuda.Cell(flexcpText, intFilaSeleccionada, intCol)
End Property
'----------------------------------------------------------------------------------------------------------------------
'Metodo: ActivarAyuda
'       Consulta:  Consulta para realizar la seleccion. Se deben enviar los datos ya filtrados.
'                        Ejemplo: Select CliCodigo, CliCodigo, Nombre from .....
'                                      - El primero es utilizado como la KEY y  desde el segundo en adelante
'                                        son los resultados a cargar en la lista.
'----------------------------------------------------------------------------------------------------------------------
Public Sub ActivarAyuda(Consulta As String, Optional AnchoForm As Currency = 6000, Optional OcultoCol1 As Boolean = True, Optional Titulo As String = "")
On Error GoTo ErrActivo
    Screen.MousePointer = 11
    
    Set RsAux = cBase.OpenResultset(Consulta, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No se encontraron datos para los filtros seleccionados.", vbInformation, "Lista de Ayuda"
    Else
        frmAyuda.bolHidePrimera = OcultoCol1
        If Titulo <> "" Then frmAyuda.Caption = Titulo
        frmAyuda.Width = AnchoForm
        frmAyuda.Show vbModal
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
ErrActivo:
    clsGeneral.OcurrioError "Ocurrio un error al iniciar el metodo ActivarAyuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub BotonAnterior()
    With Toolbar
        .Buttons("anterior").Enabled = True
        .Buttons("siguiente").Enabled = True
        .Buttons("primero").Enabled = True
    End With
    AccionAnterior
End Sub
Private Sub BotonCancelar()
    On Error Resume Next
    Unload Me
End Sub
Private Sub BotonPrimero()
    With Toolbar
        .Buttons("anterior").Enabled = True
        .Buttons("siguiente").Enabled = True
        .Buttons("primero").Enabled = True
    End With
    AccionPrimero
End Sub
Private Sub BotonSeleccionar()
    On Error Resume Next
    vsAyuda_DblClick
End Sub
Private Sub BotonSiguiente()
    With Toolbar
        .Buttons("anterior").Enabled = True
        .Buttons("siguiente").Enabled = True
        .Buttons("primero").Enabled = True
    End With
    AccionSiguiente
End Sub
Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyA: If Toolbar.Buttons("anterior").Enabled Then BotonAnterior
            Case vbKeyS: If Toolbar.Buttons("siguiente").Enabled Then BotonSiguiente
            Case vbKeyP: If Toolbar.Buttons("primero").Enabled Then BotonPrimero
            Case vbKeyX:  intFilaSeleccionada = 0: Unload Me
        End Select
    Else
        Select Case KeyCode
            Case vbKeyEscape: intFilaSeleccionada = 0: Unload Me
        End Select
    End If
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad
    
    'Inicializo variables.--------------------------------
    PosLista = 0: Cant = 0: intFilaSeleccionada = 0
    'Seteo la lista de ayuda ------------------------------
    DoyFormatoAGrilla
    '---------------------------------------------------------
    Cant = 0
    CargoDatosConsulta
    If vsAyuda.Rows > 1 Then vsAyuda.Select 1, 0, 1, vsAyuda.Cols - 1
    Exit Sub

ErrLoad:
    clsGeneral.OcurrioError "Ocurrio un error al iniciar el formulario de lista de ayuda.", Err.Description
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    'PictBotones.Top = Me.ScaleHeight - (PictBotones.Height + Status.Height + 40)
    'PictBotones.Left = 5
    'PictBotones.Width = Me.ScaleWidth - 140
    'bCancelar.Left = PictBotones.Width - (bCancelar.Width + 50)
    'bSeleccionar.Left = bCancelar.Left - (bSeleccionar.Width + 50)
    vsAyuda.Left = 5
    vsAyuda.Width = Me.ScaleWidth
    vsAyuda.Height = Me.ScaleHeight - (Toolbar.Height + Status.Height + 80)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    RsAux.Close
    Exit Sub
End Sub
Private Sub CargoDatosConsulta()
On Error GoTo ErrCDC
Dim Contador As Integer

    Screen.MousePointer = 11
    With vsAyuda
        .Rows = 1
        .Refresh
        .Redraw = False
        Do While Not RsAux.EOF And Cant < cCantidadMaxima
            With vsAyuda
                .AddItem ""
                For Contador = 0 To RsAux.rdoColumns.Count - 1
                    If Not IsNull(RsAux(Contador)) Then
                        .Cell(flexcpText, .Rows - 1, Contador) = Trim(RsAux(Contador))
                    Else
                        .Cell(flexcpText, .Rows - 1, Contador) = ""
                    End If
                Next Contador
            End With
            RsAux.MoveNext
            Cant = Cant + 1
        Loop
        .AutoSize 0, .Cols - 1
        .Redraw = True
    End With
    
    If Cant > cCantidadMaxima Then Cant = cCantidadMaxima
    
    If RsAux.EOF Then
        With Toolbar
            .Buttons("anterior").Enabled = True
            .Buttons("primero").Enabled = True
            .Buttons("siguiente").Enabled = False
        End With
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
ErrCDC:
    vsAyuda.Redraw = True
    clsGeneral.OcurrioError "Error al cargar la información en la lista.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionAnterior()
On Error GoTo ErrAA
Dim UltimaPosicion As Long
    
    Toolbar.Buttons("siguiente").Enabled = True
    If RsAux.EOF And vsAyuda.Rows > 1 And RsAux.AbsolutePosition = -1 Then
        Screen.MousePointer = 11
        RsAux.MoveLast
        Screen.MousePointer = 0
        UltimaPosicion = PosLista + (vsAyuda.Rows - 1) + 1
        RsAux.MoveNext
        Screen.MousePointer = 0
    Else
        UltimaPosicion = PosLista + Cant + 1
    End If
    If UltimaPosicion - Cant - cCantidadMaxima >= 1 Then
        RsAux.Move UltimaPosicion - Cant - cCantidadMaxima, 1
        Cant = 0
        CargoDatosConsulta
        PosLista = PosLista - Cant
        Screen.MousePointer = 0
    Else
        MsgBox "Se ha llegado al principio de la consulta.", vbInformation, "ATENCIÓN"
        With Toolbar
            .Buttons("anterior").Enabled = False
            .Buttons("primero").Enabled = True
        End With
    End If
    vsAyuda.SetFocus
    Exit Sub
ErrAA:
    clsGeneral.OcurrioError "Error inesperado al ejecutar la acción.", Err.Description
End Sub
Private Sub AccionPrimero()
On Error GoTo ErrAPSF
    Screen.MousePointer = 11
    If Not RsAux.BOF Then
        With Toolbar
            .Buttons("primero").Enabled = False
            .Buttons("anterior").Enabled = False
            .Buttons("siguiente").Enabled = True
        End With
        RsAux.MoveFirst
        Cant = 0
        PosLista = 0
        CargoDatosConsulta
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrAPSF:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error inesperado al intentar acceder al primer registro.", Err.Description
End Sub
Private Sub AccionSiguiente()
On Error GoTo ErrAS
    
    Toolbar.Buttons("primero").Enabled = True
    If Not RsAux.EOF Then
        Toolbar.Buttons("anterior").Enabled = True
        PosLista = PosLista + Cant
        Cant = 0
        CargoDatosConsulta
    Else
        MsgBox "Se llegó al final de la selección.", vbExclamation, "ATENCIÓN"
    End If
    Exit Sub
ErrAS:
    clsGeneral.OcurrioError "Error inesperado al ejecutar la acción.", Err.Description
End Sub
Private Sub DoyFormatoAGrilla()
On Error GoTo ErrDFAO
Dim strEncabezado As String, Contador As Integer
Dim arrType() As Integer

    ReDim arrType(RsAux.rdoColumns.Count - 1)
    
    strEncabezado = ""
    For Contador = 0 To RsAux.rdoColumns.Count - 1
        arrType(Contador) = RsAux.rdoColumns(Contador).Type
        strEncabezado = strEncabezado & Trim(RsAux.rdoColumns(Contador).Name) & "|"
    Next
    'Le quito el último separador.
    strEncabezado = VBA.Left(strEncabezado, Len(strEncabezado) - 1)
    
    With vsAyuda
        .Redraw = False
        .Rows = 1: .Cols = 1: .ExtendLastCol = True
        .FormatString = strEncabezado
        For Contador = 0 To RsAux.rdoColumns.Count - 1
            .ColDataType(Contador) = RetornoDataTypeGrilla(arrType(Contador))
        Next
        .ColHidden(0) = bolHidePrimera
        .Redraw = True
    End With
    Exit Sub
    
ErrDFAO:
    clsGeneral.OcurrioError "Error inesperado al intentar inicializar la grilla de ayuda.", Err.Description
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "salir": BotonCancelar
        Case "seleccionar": BotonSeleccionar
        Case "primero": BotonPrimero
        Case "anterior": BotonAnterior
        Case "siguiente": BotonSiguiente
    End Select
End Sub

Private Sub vsAyuda_DblClick()
On Error Resume Next
    If vsAyuda.Row > 0 Then intFilaSeleccionada = vsAyuda.Row: Me.Hide  'Unload Me
End Sub

Private Sub vsAyuda_GotFocus()
    Status.SimpleText = "Para elegir seleccione una fila y de doble click o presione enter."
End Sub

Private Sub vsAyuda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And vsAyuda.Row > 0 Then Call vsAyuda_DblClick
End Sub
