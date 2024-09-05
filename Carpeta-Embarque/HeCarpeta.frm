VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form HeCarpeta 
   Caption         =   "Lista de Ayuda"
   ClientHeight    =   4230
   ClientLeft      =   4320
   ClientTop       =   2355
   ClientWidth     =   4695
   FillColor       =   &H00FFFFFF&
   Icon            =   "HeCarpeta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4230
   ScaleWidth      =   4695
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PictBotones 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   4635
      TabIndex        =   1
      Top             =   3720
      Width           =   4695
      Begin VB.CheckBox chRepuestos 
         Caption         =   "Incluir Repuestos"
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "HeCarpeta.frx":0742
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Ir a la primer página.[Ctrl+P]"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Picture         =   "HeCarpeta.frx":088C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Ir a la siguiente página.[Ctrl+S]"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "HeCarpeta.frx":09D6
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Ir a la página anterior.[Ctrl+A]"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSeleccionar 
         Height          =   310
         Left            =   3840
         Picture         =   "HeCarpeta.frx":0B20
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Seleccionar.[Enter]"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "HeCarpeta.frx":0C6A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Consultar.[Ctrl+E]"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   4200
         Picture         =   "HeCarpeta.frx":0DB4
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Cancelar.[Ctrl+X]"
         Top             =   120
         Width           =   310
      End
   End
   Begin MSComctlLib.ListView lAyuda 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6376
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "HeCarpeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const cCantidadMaxima = 100
Private PosLista As Long, Cant As Long
Private sConsulta As String                     'Consulta a resolver
Private iSeleccionado As Long
Public Property Get pConsulta() As String
    pConsulta = ""
End Property
Public Property Let pConsulta(Texto As String)
    sConsulta = Texto
End Property
Public Property Get pSeleccionado() As Long
    pSeleccionado = iSeleccionado
End Property
Public Property Let pSeleccionado(Codigo As Long)
    iSeleccionado = Codigo
End Property
Private Sub bAnterior_Click()
    bSiguiente.Enabled = True
    bAnterior.Enabled = True: bPrimero.Enabled = True
    AccionAnterior
End Sub
Private Sub bCancelar_Click()
    On Error Resume Next
    Unload Me
End Sub
Private Sub bConsultar_Click()
    If InicializoConsulta Then CargoDatosConsulta
End Sub
Private Sub bPrimero_Click()
    bSiguiente.Enabled = True
    bPrimero.Enabled = True
    bAnterior.Enabled = True
    AccionPrimero
End Sub
Private Sub bSeleccionar_Click()
    On Error Resume Next
    lAyuda_DblClick
End Sub
Private Sub bSiguiente_Click()
    bSiguiente.Enabled = True
    bPrimero.Enabled = True
    bAnterior.Enabled = True
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
        Case vbKeyE
            bConsultar_Click
        Case vbKeyP
            bPrimero_Click
        Case vbKeyS
            bSiguiente_Click
        Case vbKeyA
            bAnterior_Click
        Case vbKeyX
            iSeleccionado = 0: Unload Me
        End Select
    Else
        Select Case KeyCode
            Case vbKeyEscape
                iSeleccionado = 0: Unload Me
            Case vbKeyReturn
                lAyuda_DblClick
        End Select
    End If

    If KeyCode = vbKeyEscape Then Unload Me
End Sub
Private Sub Form_Load()
    On Error GoTo ErrLoad
    
    SetearLView lvValores.UnClickIcono Or lvValores.FullRow Or lvValores.HeaderDragDrop, lAyuda
    
    'Picture que contiene los botones
    PictBotones.BorderStyle = 0
    'Inicializo variables.--------------------------------
    PosLista = 0: Cant = 0: iSeleccionado = 0
        
    
    'Seteo la lista de ayuda ------------------------------
    DoyFormatoAObjetos
    CargoDatosConsulta
    If lAyuda.ListItems.Count > 0 Then Set lAyuda.SelectedItem = lAyuda.ListItems(1)
    '---------------------------------------------------------
    
    'Anulo los botones de inicio.----------------------
    bAnterior.Enabled = False: bPrimero.Enabled = False

    Exit Sub

ErrLoad:
    MsgBox "Ocurrió un error al cargar los datos en la lista. Verifique el ingreso de caracteres no válidos.", vbExclamation, "ATENCIÓN"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
On Error GoTo ErrResize
    
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.ScaleHeight < 1000 Then Exit Sub
    Screen.MousePointer = 11
    PictBotones.Top = Me.ScaleHeight - PictBotones.Height        'La escala menos el largo del status me da el status.top
    PictBotones.Width = Me.ScaleWidth - 120
    lAyuda.Left = 5
    lAyuda.Width = Me.ScaleWidth
    lAyuda.Height = Me.ScaleHeight - PictBotones.Height
    Refresh
    Screen.MousePointer = 0
    Exit Sub
ErrResize:
    Screen.MousePointer = 0
    MsgBox "Ocurrio un error inesperado." & vbCr & "Error: " & Trim(Err.Description), vbExclamation, "ATENCIÓN"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrME
    Screen.MousePointer = 11
    On Error Resume Next
    CierroCursor
    Screen.MousePointer = 0
    Exit Sub
ErrME:
    Screen.MousePointer = 0
    MsgBox "Ocurrio un error al cerrar la lista de ayuda. " & Trim(Err.Description), vbCritical, "ATENCIÓN"
    On Error Resume Next
    CierroCursor
End Sub
Private Sub lAyuda_DblClick()
On Error Resume Next
    If lAyuda.ListItems.Count > 0 Then
        Screen.MousePointer = 11
        iSeleccionado = CLng(Mid(lAyuda.SelectedItem.Key, 2, Len(lAyuda.SelectedItem.Key)))
        Unload Me
        Screen.MousePointer = 0
    End If
End Sub

Private Sub CargoDatosConsulta()
On Error GoTo ErrCDC
Dim CodCarpeta As Long

    Screen.MousePointer = 11
    CodCarpeta = 0
    Do While Not RsHelp.EOF And Cant <= 100
        If CodCarpeta <> RsHelp(0) Then
            Set itmx = lAyuda.ListItems.Add(, "A" & RsHelp(0), Trim(CStr(RsHelp(1))))
            CodCarpeta = RsHelp(0) 'Codigo
            itmx.SubItems(1) = Format(RsHelp(2), "d-mmm yy")
            itmx.SubItems(2) = Format(RsHelp(3), "d-mmm yy")
            itmx.SubItems(3) = Trim(CStr(RsHelp(4)))
            itmx.SubItems(4) = Trim(CStr(RsHelp(5)))
        Else
            If InStr(itmx.SubItems(4), RsHelp(5)) = 0 Then
                itmx.SubItems(4) = itmx.SubItems(4) + ", " + Trim(RsHelp(5))
            End If
        End If
        RsHelp.MoveNext
        Cant = Cant + 1
    Loop
    If RsHelp.EOF Then
        bSiguiente.Enabled = False
        bAnterior.Enabled = True
        bPrimero.Enabled = True
    End If
    If Cant > cCantidadMaxima Then Cant = cCantidadMaxima
    If lAyuda.ListItems.Count > 0 Then Set lAyuda.SelectedItem = lAyuda.ListItems(1): AutoSizeColumns lAyuda
    Screen.MousePointer = 0
    Exit Sub
    
ErrCDC:
    MsgBox "Ocurrio un error al cargar la lista de ayuda.", vbExclamation, "ATENCIÓN"
    Screen.MousePointer = 0
End Sub

Private Sub lAyuda_KeyDown(KeyCode As Integer, Shift As Integer)
    If lAyuda.ListItems.Count = 0 Then Exit Sub
    Select Case KeyCode
        Case vbKeyAdd
            If bSiguiente.Enabled Then bSiguiente_Click
        Case vbKeySubtract
            If bAnterior.Enabled Then bAnterior_Click
        Case vbKeyMultiply
            If bPrimero.Enabled Then bPrimero_Click
    End Select
End Sub

Private Sub lAyuda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And lAyuda.ListItems.Count > 0 Then Call lAyuda_DblClick
End Sub
Private Sub AccionAnterior()
On Error GoTo ErrAA
Dim UltimaPosicion As Long
    
    bSiguiente.Enabled = True
    If RsHelp.EOF And lAyuda.ListItems.Count > 0 And RsHelp.AbsolutePosition = -1 Then
        Screen.MousePointer = 11
        RsHelp.MoveLast
        Screen.MousePointer = 0
        UltimaPosicion = PosLista + lAyuda.ListItems.Count + 1
        RsHelp.MoveNext
        Screen.MousePointer = 0
    Else
        UltimaPosicion = PosLista + Cant + 1
    End If
    If UltimaPosicion - Cant - cCantidadMaxima >= 1 Then
        RsHelp.Move UltimaPosicion - Cant - cCantidadMaxima, 1
        Cant = 0
        CargoDatosConsulta
        PosLista = PosLista - Cant
        Screen.MousePointer = 0
    Else
        MsgBox "Se ha llegado al principio de la consulta.", vbInformation, "ATENCIÓN"
        bPrimero.Enabled = False: bAnterior.Enabled = False
    End If
    lAyuda.SetFocus
    Exit Sub
ErrAA:
'    clsGeneral.OcurrioError "Ocurrio un error inesperado."
End Sub
Private Sub AccionPrimero()
On Error GoTo ErrAPSF
    Screen.MousePointer = 11
    If Not RsHelp.BOF Then
        bPrimero.Enabled = False:
        bAnterior.Enabled = False
        bSiguiente.Enabled = True
        RsHelp.MoveFirst
        Cant = 0
        PosLista = 0
        CargoDatosConsulta
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrAPSF:
    Screen.MousePointer = 0
'    clsGeneral.OcurrioError "Ocurrio un error al acceder al primer registro."
End Sub
Private Sub AccionSiguiente()
On Error GoTo ErrAS
    bPrimero.Enabled = True
    If Not RsHelp.EOF Then
        bAnterior.Enabled = True
        PosLista = PosLista + Cant
        Cant = 0
        CargoDatosConsulta
    Else
        MsgBox "Se llegó al final de la selección.", vbExclamation, "ATENCIÓN"
    End If
    Exit Sub
ErrAS:
'    clsGeneral.OcurrioError "Ocurrio un error inesperado."
End Sub
Private Sub DoyFormatoAObjetos()
On Error GoTo ErrDFAO
Dim AuxAlineacion  As Integer, Contador As Integer
    
    Screen.MousePointer = 11
    For Contador = 1 To RsHelp.rdoColumns.Count - 1
        
        'Determino la alineación:
            'Si es texto va a la izquierda y si es número va a la derecha.
        If Contador = 1 Then
            AuxAlineacion = lvwColumnLeft
        Else
            Select Case RsHelp.rdoColumns(Contador).Type
                Case rdTypeCHAR, rdTypeLONGVARCHAR, rdTypeVARCHAR
                    AuxAlineacion = lvwColumnLeft
                Case Else
                    AuxAlineacion = lvwColumnRight
            End Select
        End If
        lAyuda.ColumnHeaders.Add , , Trim(RsHelp.rdoColumns(Contador).Name), 800, AuxAlineacion
    Next
    Screen.MousePointer = 0
    Exit Sub

ErrDFAO:
    Screen.MousePointer = 0
'    clsGeneral.OcurrioError "Ocurrio un error al dar formato a la lista.", Trim(Err.Description)
End Sub
Private Function InicializoConsulta() As Boolean
On Error GoTo ErrIC
    
    Screen.MousePointer = 11
    'Limpio la lista.---------------------------------------------
    lAyuda.ListItems.Clear
'    lAyuda.ColumnHeaders.Clear
    '--------------------------------------------------------------
    
    'Inicializo contadores de la lista.---------------------
    Cant = 0: PosLista = 0
    '--------------------------------------------------------------
    
    'Pido la consulta.------------------------------------------
    CierroCursor
    
    If Not chRepuestos.Value Then
        Cons = sConsulta & " And AFoArticulo NOT IN(Select AGrArticulo From ArticuloGrupo" _
                            & " Where AGrGrupo = " & paRepuesto & ")"
    Else
        Cons = sConsulta
    End If
    Cons = Cons & " Order by 2"

    If PidoConsulta(Cons) = 0 Then InicializoConsulta = False: Exit Function
    '--------------------------------------------------------------
    InicializoConsulta = True
    Screen.MousePointer = 0
    Exit Function
ErrIC:
    InicializoConsulta = False
    Screen.MousePointer = 0
End Function
'La utilidad de esta función es para no activar el load...............
Public Function PidoConsulta(strConsulta) As Integer
On Error GoTo ErrPC
    Screen.MousePointer = 11
    sConsulta = strConsulta
    Set RsHelp = cBase.OpenResultset(strConsulta, rdOpenDynamic, rdConcurReadOnly)
    If RsHelp.EOF Then PidoConsulta = 1 Else PidoConsulta = 2
    Screen.MousePointer = 0
    Exit Function
ErrPC:
    PidoConsulta = 0
    Screen.MousePointer = 0
    MsgBox "Ocurrio un error al consultar." & vbCr & "Error: " & Trim(Err.Description), vbCritical, "ATENCIÓN"
End Function

Public Sub CierroCursor()
On Error Resume Next
    RsHelp.Close
End Sub
