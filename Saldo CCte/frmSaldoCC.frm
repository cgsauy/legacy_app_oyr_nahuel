VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmSaldoCC 
   Caption         =   "Saldo Cuentas Corrientes"
   ClientHeight    =   4770
   ClientLeft      =   3015
   ClientTop       =   3450
   ClientWidth     =   7860
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSaldoCC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   7860
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   9
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir del formulario"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   300
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Caption         =   ""
            Key             =   "modificar"
            Description     =   ""
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
      Begin ComctlLib.StatusBar Status1 
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   1500
         Visible         =   0   'False
         Width           =   9180
         _ExtentX        =   16193
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
               Object.Width           =   8440
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Saldo Al"
      ForeColor       =   &H00000080&
      Height          =   1275
      Left            =   60
      TabIndex        =   14
      Top             =   480
      Width           =   7695
      Begin VB.TextBox tProveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1140
         MaxLength       =   40
         TabIndex        =   1
         Top             =   300
         Width           =   4455
      End
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   6300
         TabIndex        =   3
         Top             =   300
         Width           =   1275
      End
      Begin VB.TextBox tSPeso 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1140
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox tSDolar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   4140
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox tComentario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   9
         Top             =   900
         Width           =   6435
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Proveedor:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha:"
         Height          =   255
         Left            =   5700
         TabIndex        =   2
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo &Pesos:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   630
         Width           =   915
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo &Dólares:"
         Height          =   255
         Left            =   3060
         TabIndex        =   6
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Comentarios:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   900
         Width           =   975
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsGrilla 
      Height          =   2535
      Left            =   60
      TabIndex        =   10
      Top             =   1800
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4471
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   4
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
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   4515
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   5741
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6660
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSaldoCC.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSaldoCC.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSaldoCC.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSaldoCC.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSaldoCC.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSaldoCC.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSaldoCC.frx":0AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSaldoCC.frx":0C88
            Key             =   ""
         EndProperty
      EndProperty
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
   End
   Begin VB.Menu MnuSalirDelFormulario 
      Caption         =   "Salir"
      Begin VB.Menu MnuSalir 
         Caption         =   "&Del Formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmSaldoCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sNuevo As Boolean, sModificar As Boolean


Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()

    On Error GoTo ErrLoad
    ObtengoSeteoForm Me
    sNuevo = False: sModificar = False
    
    OcultoCampos
    
    InicializoGrilla
    Botones False, False, False, False, False, Toolbar1, Me
    Screen.MousePointer = 0
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al cargar el formulario.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    Frame1.Width = Me.ScaleWidth - (Frame1.Left * 2)
    
    vsGrilla.Width = Me.ScaleWidth - (vsGrilla.Left * 2)
    vsGrilla.Height = Me.ScaleHeight - (vsGrilla.Top + Status.Height + 70)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    GuardoSeteoForm Me
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
End Sub

Private Sub Label1_Click()
    Foco tProveedor
End Sub
Private Sub Label2_Click()
    Foco tFecha
End Sub
Private Sub Label3_Click()
    Foco tSPeso
End Sub
Private Sub Label4_Click()
    Foco tSDolar
End Sub
Private Sub Label5_Click()
    Foco tComentario
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuEliminar_Click()
    AccionEliminar
End Sub

Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuModificar_Click()
    AccionModificar
End Sub

Private Sub MnuNuevo_Click()
    AccionNuevo
End Sub

Private Sub MnuSalir_Click()
    Unload Me
End Sub

Private Sub tComentario_GotFocus()
    With tComentario
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub
Private Sub tComentario_LostFocus()
    tComentario.SelStart = 0
End Sub
Private Sub tFecha_GotFocus()
    With tFecha
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy") Else tFecha.Text = ""
        Foco tSPeso
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        Case "salir": Unload Me
    End Select
    
End Sub

Private Sub tProveedor_Change()
    If Val(tProveedor.Tag) <> 0 Then
        tProveedor.Tag = 0: vsGrilla.Rows = 1: Botones False, False, False, False, False, Toolbar1, Me
    End If
End Sub

Private Sub tProveedor_GotFocus()
    With tProveedor: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tProveedor_KeyPress(KeyAscii As Integer)
On Error GoTo errBuscar
    
    If KeyAscii = vbKeyReturn Then
        If Val(tProveedor.Tag) <> 0 Or Trim(tProveedor.Text) = "" Then Foco tFecha: Exit Sub
        Screen.MousePointer = 11
        Dim aQ As Long, aIdProveedor As Long, aTexto As String
        
        aQ = 0
        Cons = "Select PClCodigo, PClFantasia, PClNombre from ProveedorCliente " _
                & " Where PClNombre like '" & Trim(tProveedor.Text) & "%' Or PClFantasia like '" & Trim(tProveedor.Text) & "%'"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            aQ = 1: aIdProveedor = RsAux!PClCodigo: aTexto = Trim(RsAux!PClFantasia)
            RsAux.MoveNext: If Not RsAux.EOF Then aQ = 2
        End If
        RsAux.Close
        
        Select Case aQ
            Case 0: MsgBox "No existe una empresa para el con el nombre ingresado.", vbExclamation, "No existe Empresa"
            
            Case 1:
                    tProveedor.Text = aTexto: tProveedor.Tag = aIdProveedor
                    CargoGrilla
                    Foco tFecha
        
            Case 2:
                    Dim aLista As New clsListadeAyuda
                    aLista.ActivoListaAyuda Cons, False, txtConexion, 5500
                    If aLista.ValorSeleccionado <> 0 Then
                        tProveedor.Text = Trim(aLista.ItemSeleccionado)
                        tProveedor.Tag = aLista.ValorSeleccionado
                        CargoGrilla
                        Foco tFecha
                    Else
                        tProveedor.Text = ""
                    End If
                    Set aLista = Nothing
        End Select
    End If
    Screen.MousePointer = 0
    Exit Sub

errBuscar:
    clsGeneral.OcurrioError "Ocurrió un error al procesar la lista de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub InicializoGrilla()
    
    With vsGrilla
        .Redraw = False
        .ExtendLastCol = True
        .Editable = True
        .Rows = 1: .FixedCols = 0: .Cols = 1
        .FormatString = "Fecha|>Saldo $|>Saldo U$S|Comentarios"
        .ColWidth(0) = 950: .ColWidth(1) = 1200: .ColWidth(2) = 1100
        
        .AllowUserResizing = flexResizeColumns
        .HighLight = flexHighlightAlways
        .Redraw = True
    End With
    
End Sub

Private Sub OcultoCampos()
    
    tProveedor.Enabled = True: tProveedor.BackColor = Colores.Blanco
    
    tFecha.Text = "": tFecha.Enabled = False: tFecha.BackColor = Colores.Inactivo
    tSPeso.Text = "": tSPeso.Enabled = False: tSPeso.BackColor = Colores.Inactivo
    tSDolar.Text = "": tSDolar.Enabled = False: tSDolar.BackColor = Colores.Inactivo
    tComentario.Text = "": tComentario.Enabled = False: tComentario.BackColor = Colores.Inactivo
    
    vsGrilla.Enabled = True
    
End Sub

Private Sub HabilitoCampos()

    tProveedor.Enabled = False: tProveedor.BackColor = Colores.Inactivo
    
    tFecha.Text = "": tFecha.Enabled = True: tFecha.BackColor = Colores.Blanco
    tSPeso.Text = "": tSPeso.Enabled = True: tSPeso.BackColor = Colores.Blanco
    tSDolar.Text = "": tSDolar.Enabled = True: tSDolar.BackColor = Colores.Blanco
    tComentario.Text = "": tComentario.Enabled = True: tComentario.BackColor = Colores.Blanco
    
    vsGrilla.Enabled = False
    
End Sub

Private Sub AccionNuevo()

    If Val(tProveedor.Tag) = 0 Then
        Botones False, False, False, False, False, Toolbar1, Me
        MsgBox "Debe seleccionar un proveedor para ingresar el saldo.", vbInformation, "Seleccione el Proveedor"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    sNuevo = True
    Botones False, False, False, True, True, Toolbar1, Me
    HabilitoCampos
    
    tFecha.Text = Format(Date, "dd/mm/yyyy"): Foco tFecha
    Screen.MousePointer = 0
    
End Sub

Private Sub AccionModificar()
    
    If Val(tProveedor.Tag) = 0 Then
        Botones False, False, False, False, False, Toolbar1, Me
        MsgBox "Debe seleccionar un proveedor para ingresar el saldo.", vbInformation, "Seleccione el Proveedor"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    sModificar = True
    HabilitoCampos
    Botones False, False, False, True, True, Toolbar1, Me
    
    tFecha.Text = Trim(vsGrilla.Cell(flexcpText, vsGrilla.Row, 0))
    tSPeso.Text = Trim(vsGrilla.Cell(flexcpText, vsGrilla.Row, 1))
    tSDolar.Text = Trim(vsGrilla.Cell(flexcpText, vsGrilla.Row, 2))
    tComentario.Text = Trim(vsGrilla.Cell(flexcpText, vsGrilla.Row, 3))
    Foco tFecha
    Screen.MousePointer = 0
    
End Sub
Private Sub AccionEliminar()
    
    On Error GoTo ErrAE

    If MsgBox("Confirma eliminar el saldo ingresado para el  " & Trim(vsGrilla.Cell(flexcpText, vsGrilla.Row, 0)), vbQuestion + vbYesNo, "ELIMINAR") = vbNo Then Exit Sub
    
    Screen.MousePointer = 0
    
    Cons = "Select * From SaldoCCte " & _
                    " Where SCCProveedor = " & Val(tProveedor.Tag) & _
                    " And SCCFecha = '" & Format(vsGrilla.Cell(flexcpText, vsGrilla.Row, 0), sqlFormatoFH) & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "Otra terminal pudo eliminar el saldo seleccionado, verifique.", vbExclamation, "ATENCIÓN"
    Else
        RsAux.Delete
        RsAux.Close
    End If
    
    CargoGrilla
    Screen.MousePointer = 0
    Exit Sub
    
ErrAE:
    clsGeneral.OcurrioError "Ocurrio un error al intentar eliminar el saldo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionCancelar()
    sNuevo = False: sModificar = False
    OcultoCampos
    CargoGrilla
End Sub

Private Sub AccionGrabar()
    
    If Not ValidoDatos Then Exit Sub
    On Error GoTo errGrabar
    If MsgBox("Confirma almacenar los datos ingresados", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    Screen.MousePointer = 11
            
    If sNuevo Then
        Cons = "Select * From SaldoCCte Where SCCProveedor = " & Val(tProveedor.Tag)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.AddNew
        CargoCamposBD
        RsAux.Update: RsAux.Close
        
    Else
    
        Cons = "Select * From SaldoCCte " & _
                    " Where SCCProveedor = " & Val(tProveedor.Tag) & _
                    " And SCCFecha = '" & Format(vsGrilla.Cell(flexcpText, vsGrilla.Row, 0), sqlFormatoFH) & "'"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.Edit
        CargoCamposBD
        RsAux.Update: RsAux.Close
    
    End If
    
    On Error Resume Next
    AccionCancelar
    Screen.MousePointer = 0
    Exit Sub
    
errGrabar:
    clsGeneral.OcurrioError "Ocurrió un error al grabar los datos. Verifique si no existe un saldo para la fecha ingresada.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub CargoCamposBD()
    
    RsAux!SCCProveedor = Val(tProveedor.Tag)
    RsAux!SCCFecha = Format(tFecha.Text, sqlFormatoFH)
    RsAux!SCCSaldoP = CCur(tSPeso.Text)
    RsAux!SCCSaldoD = CCur(tSDolar.Text)
    If Trim(tComentario.Text) = "" Then RsAux!SCCComentario = Null Else RsAux!SCCComentario = Trim(tComentario.Text)
    RsAux!SCCUsuario = paCodigoDeUsuario
    
End Sub


Private Sub CargoGrilla()

On Error GoTo ErrCG

    Screen.MousePointer = 11
    vsGrilla.Rows = 1
    If Val(tProveedor.Tag) = 0 Then Screen.MousePointer = 0: Botones False, False, False, False, False, Toolbar1, Me: Exit Sub
    
    Cons = "Select * From SaldoCCte " & _
            " Where SCCProveedor = " & Val(tProveedor.Tag) & _
            " Order by SCCFecha Desc"
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        With vsGrilla
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!SCCFecha, "dd/mm/yyyy")
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!SCCSaldoP, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!SCCSaldoD, FormatoMonedaP)
            If Not IsNull(RsAux!SCCComentario) Then .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!SCCComentario)
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If vsGrilla.Rows > 1 Then
        vsGrilla.Select 1, 0, 1, vsGrilla.Cols - 1
        Botones True, True, True, False, False, Toolbar1, Me
    Else
        Botones True, False, False, False, False, Toolbar1, Me
    End If
    Screen.MousePointer = 0
    Exit Sub
    
ErrCG:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los saldos ya grabados.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function ValidoDatos() As Boolean
    
    On Error GoTo errValido
    ValidoDatos = False
    
    If Val(tProveedor.Tag) = 0 Then
        MsgBox "Debe seleccionar un proveedor para grabar el saldo.", vbExclamation, "Posible Error"
        Exit Function
    End If
    
    If Not IsDate(tFecha.Text) Then
        MsgBox "La fecha ingresada para el saldo no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tFecha: Exit Function
    End If
    
    If Not IsNumeric(tSPeso.Text) Then tSPeso.Text = "0.00"
    If Not IsNumeric(tSDolar.Text) Then tSDolar.Text = "0.00"
    
    ValidoDatos = True
    Exit Function

errValido:
End Function


Private Sub tSDolar_GotFocus()
    With tSDolar: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tSDolar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tSDolar.Text) Then tSDolar.Text = Format(tSDolar.Text, FormatoMonedaP) Else tSDolar.Text = "0.00"
        Foco tComentario
    End If
End Sub

Private Sub tSPeso_GotFocus()
    With tSPeso: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tSPeso_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tSPeso.Text) Then tSPeso.Text = Format(tSPeso.Text, FormatoMonedaP) Else tSPeso.Text = "0.00"
        Foco tSDolar
    End If
End Sub
