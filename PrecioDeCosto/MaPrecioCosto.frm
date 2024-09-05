VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form MaPrecioCosto 
   Caption         =   "Precios de Costo"
   ClientHeight    =   4665
   ClientLeft      =   2895
   ClientTop       =   3060
   ClientWidth     =   7590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MaPrecioCosto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7590
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
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
            Enabled         =   0   'False
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ficha"
      ForeColor       =   &H00000080&
      Height          =   1935
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   7335
      Begin VB.TextBox tNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   840
         MaxLength       =   30
         TabIndex        =   1
         Top             =   240
         Width           =   6375
      End
      Begin VB.TextBox tCosto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3480
         MaxLength       =   14
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox tComentario 
         Appearance      =   0  'Flat
         Height          =   555
         Left            =   120
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1200
         Width           =   7095
      End
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         MaxLength       =   14
         TabIndex        =   3
         Text            =   "00/00/2000"
         Top             =   600
         Width           =   975
      End
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   2580
         TabIndex        =   5
         Top             =   600
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
      Begin AACombo99.AACombo cIncoterm 
         Height          =   315
         Left            =   5040
         TabIndex        =   8
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
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
      Begin AACombo99.AACombo cITexto 
         Height          =   315
         Left            =   6000
         TabIndex        =   9
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.Label Label4 
         Caption         =   "&Inco.:"
         Height          =   255
         Left            =   4560
         TabIndex        =   7
         Top             =   620
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Costo:"
         Height          =   255
         Left            =   2040
         TabIndex        =   4
         Top             =   620
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   620
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Co&mentarios:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsGrilla 
      Height          =   1815
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3201
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
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   4
      GridLinesFixed  =   2
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
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   4410
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3720
      Top             =   240
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
            Picture         =   "MaPrecioCosto.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaPrecioCosto.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaPrecioCosto.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaPrecioCosto.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaPrecioCosto.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaPrecioCosto.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaPrecioCosto.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaPrecioCosto.frx":0DC8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuNuevo 
         Caption         =   "&Nuevo"
         Enabled         =   0   'False
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
      Begin VB.Menu MnuOpL1 
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
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "&Del Formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "MaPrecioCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sNuevo As Boolean, sModificar As Boolean

Private Sub cIncoterm_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cITexto
End Sub

Private Sub cITexto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tComentario
End Sub

Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tCosto
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    DoEvents
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad

    ObtengoSeteoForm Me, 1000, 500, 3840, 4230
    sNuevo = False: sModificar = False
    LimpioGrilla
    OcultoCampos
    FechaDelServidor
    'Cargo datos en los combos------------------------------------------------------------
    Cons = "Select MonCodigo, MonSigno from Moneda Order by MonSigno"
    CargoCombo Cons, cMoneda
    Cons = "Select IncCodigo, IncNombre from Incoterm Order by IncNombre"
    CargoCombo Cons, cIncoterm
    Cons = "Select CiuCodigo, CiuNombre from Ciudad Order by CiuNombre"
    CargoCombo Cons, cITexto
    '-------------------------------------------------------------------------------------------
    
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al iniciar el formulario.", Trim(Err.Description)
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    'If Me.Height < 4455 Then Me.Height = 4455
    'If Me.Width < 4770 Or Me.Width > 4770 Then Me.Width = 4770
    vsGrilla.Width = Me.ScaleWidth - (vsGrilla.Left * 2)
    vsGrilla.Height = Me.ScaleHeight - (vsGrilla.Top + 70 + Status.Height)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    GuardoSeteoForm Me
    CierroConexion
    Set clsGeneral = Nothing
    Set miconexion = Nothing
    End
End Sub
Private Sub Label1_Click()
    Foco tNombre
End Sub
Private Sub Label2_Click()
    Foco cMoneda
End Sub

Private Sub Label4_Click()
    Foco cIncoterm
End Sub

Private Sub Label5_Click()
    Foco tFecha
End Sub

Private Sub Label6_Click()
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
Private Sub MnuVolver_Click()
    Unload Me
End Sub
Private Sub AccionNuevo()
   
    sNuevo = True
    Botones False, False, False, True, True, Toolbar1, Me
    HabilitoCampos
    tFecha.Text = Format(gFechaServidor, "dd/mm/yyyy")
    Foco tFecha
    
End Sub

Private Sub AccionModificar()
    
    Botones False, False, False, True, True, Toolbar1, Me
    HabilitoCampos
    
    With vsGrilla
        tFecha.Text = Format(.Cell(flexcpText, .Row, 0), "dd/mm/yyyy")
        If .Cell(flexcpData, .Row, 1) <> 0 Then BuscoCodigoEnCombo cMoneda, .Cell(flexcpData, .Row, 1)
        tCosto.Text = Trim(.Cell(flexcpText, .Row, 2))
        If .Cell(flexcpData, .Row, 3) <> 0 Then BuscoCodigoEnCombo cIncoterm, .Cell(flexcpData, .Row, 3)
        cITexto.Text = Trim(.Cell(flexcpText, .Row, 5))
        If Trim(.Cell(flexcpText, .Row, 4)) <> "" Then tComentario.Text = Trim(.Cell(flexcpText, .Row, 4))
    End With
    
    Foco tFecha
    
End Sub

Private Sub AccionGrabar()
    
    On Error GoTo errGrabar
    
    If Not ValidoCampos Then Exit Sub
    
    If MsgBox("Confirma grabar los datos ingresados en la ficha", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    
    If sNuevo Then
    
        'Cons = "Select * From PrecioDeCosto Where PCoArticulo = " & Val(tNombre.Tag)
        'Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        'RsAux.AddNew
        'CargoCamposBD
        'RsAux.Update: RsAux.Close
        Cons = "Insert Into PrecioDeCosto (PCoArticulo, PCoFecha, PCoImporte, PCoMoneda, PCoIncoterm, PCoComentario, PCoITexto)" & _
            " Values (" & Val(tNombre.Tag) & ", '" & Format(tFecha.Text & " " & Time, "mm/dd/yyyy hh:mm:ss") & "', " & _
            CCur(tCosto.Text) & ", " & cMoneda.ItemData(cMoneda.ListIndex) & ", "
        
        If cIncoterm.ListIndex > -1 Then
            Cons = Cons & cIncoterm.ItemData(cIncoterm.ListIndex)
        Else
            Cons = Cons & " Null "
        End If
        
        If Trim(tComentario.Text) <> "" Then
            Cons = Cons & ",'" & Trim(tComentario.Text) & "'"
        Else
            Cons = Cons & ", Null"
        End If
        
        If Trim(cITexto.Text) <> "" Then
            Cons = Cons & ",'" & Trim(cITexto.Text) & "')"
        Else
            Cons = Cons & ", Null)"
        End If
        
        cBase.Execute Cons
        
    Else
        'Cons = " Select * From PrecioDeCosto " _
                & " Where PCoArticulo = " & Val(tNombre.Tag) _
                & " And PCoFecha = '" & Format(vsGrilla.Cell(flexcpText, vsGrilla.Row, 0), sqlFormatoFH) & "'"
            
        'Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        'If Not RsAux.EOF Then
        '    RsAux.Edit
        '    CargoCamposBD
        '    RsAux.Update: RsAux.Close
        'Else
        '    RsAux.Close
        '    MsgBox "El precio de costo fue eliminado, verifique.", vbExclamation, "ATENCIÓN"
        'End If
        
        Cons = "Update PrecioDeCosto " & _
                    " Set PCoFecha = '" & Format(tFecha.Text & " " & Time, "mm/dd/yyyy hh:mm:ss") & "', " & _
                    " PCoImporte = " & CCur(Format(tCosto.Text, "#,##0.00")) & "," & _
                    " PCoMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)

        If cIncoterm.ListIndex > -1 Then
            Cons = Cons & ", PCoIncoterm = " & cIncoterm.ItemData(cIncoterm.ListIndex)
        Else
            Cons = Cons & ", PCoIncoterm = Null "
        End If
        
        If Trim(tComentario.Text) <> "" Then
            Cons = Cons & ", PCoComentario = '" & Trim(tComentario.Text) & "'"
        Else
            Cons = Cons & ", PCoComentario = Null"
        End If
        
        If Trim(cITexto.Text) <> "" Then
            Cons = Cons & ", PCoITexto = '" & Trim(cITexto.Text) & "'"
        Else
            Cons = Cons & ", PCoITexto = Null"
        End If
        
        Cons = Cons & " Where PCoArticulo = " & Val(tNombre.Tag) _
                           & " And PCoFecha = '" & Format(vsGrilla.Cell(flexcpText, vsGrilla.Row, 0), sqlFormatoFH) & "'"
        
        cBase.Execute (Cons)

    End If
    
    On Error Resume Next
    CargoGrilla
    vsGrilla.SetFocus
    sModificar = False: sNuevo = False
    Screen.MousePointer = 0
    Exit Sub
    
errGrabar:
    clsGeneral.OcurrioError "Ocurrió un error al grabar los datos de la ficha.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Function ValidoCampos() As Boolean

    ValidoCampos = False
    
    If Not IsDate(tFecha.Text) Then
        MsgBox "La fecha ingresada para al costo del artículo no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tFecha: Exit Function
    End If
    
    If cMoneda.ListIndex = -1 Then
        MsgBox "No se ha seleccionado una moneda para almacenar el costo del artículo.", vbExclamation, "ATENCIÓN"
        Foco cMoneda: Exit Function
    End If
    
    If Not IsNumeric(tCosto.Text) Then
        MsgBox "El costo del artículo no es correcto. Verifique.", vbExclamation, "ATENCIÓN"
        Foco tCosto: Exit Function
    End If
    
    If Not clsGeneral.TextoValido(tComentario.Text) Then
        MsgBox "Existen comillas simples en el comentario, eliminelas.", vbExclamation, "ATENCIÓN"
        Foco tComentario: Exit Function
    End If
    ValidoCampos = True
    
End Function

Private Sub AccionEliminar()
    
    On Error GoTo ErrAE
    
    If MsgBox("¿Confirma eliminar el precio para la fecha '" & Trim(vsGrilla.Cell(flexcpText, vsGrilla.Row, 0)) & "' ?", vbQuestion + vbYesNo, "ELIMINAR") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    Cons = "Delete PrecioDeCosto " _
            & " Where PCoArticulo = " & Val(tNombre.Tag) _
            & " And PCoFecha = '" & Format(vsGrilla.Cell(flexcpText, vsGrilla.Row, 0), sqlFormatoFH) & "'"
    cBase.Execute (Cons)
    CargoGrilla
    Screen.MousePointer = 0
    Exit Sub
ErrAE:
    clsGeneral.OcurrioError "Ocurrió un error al eliminar el registro.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Sub AccionCancelar()
    vsGrilla.Enabled = True
    sNuevo = False: sModificar = False
    CargoGrilla
End Sub

Private Sub tCosto_GotFocus()
    With tCosto: .SelStart = 0: .SelLength = Len(.Text): End With
    Status.SimpleText = "Ingrese el costo del artículo."
End Sub

Private Sub tCosto_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tCosto.Text) Then tCosto.Text = Format(tCosto.Text, FormatoMonedaP) Else tCosto.Text = ""
        Foco cIncoterm
    End If
    
End Sub

Private Sub tCosto_LostFocus()
    Status.SimpleText = ""
End Sub


Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cMoneda
End Sub

Private Sub tFecha_LostFocus()
    If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy")
End Sub

Private Sub tNombre_Change()
    If Val(tNombre.Tag) > 0 Then
        LimpioGrilla
        Botones False, False, False, False, False, Toolbar1, Me
        tNombre.Tag = ""
    End If
End Sub

Private Sub tNombre_GotFocus()
    With tNombre
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Ingrese un artículo."
End Sub

Private Sub tNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(tNombre.Text) <> "" Then
        If Val(tNombre.Tag) = 0 Then
            If IsNumeric(tNombre.Text) Then
                BuscoArticuloPorCodigo CLng(tNombre.Text)
            Else
                BuscoArticuloPorNombre
                If tNombre.Tag = "" Then
                    LimpioGrilla
                    Botones False, False, False, False, False, Toolbar1, Me
                End If
            End If
        End If
    End If
End Sub

Private Sub tNombre_LostFocus()
    Status.SimpleText = ""
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

Private Sub OcultoCampos()
    
    tNombre.BackColor = vbWhite: tNombre.Enabled = True
    tFecha.BackColor = Colores.Gris: tFecha.Enabled = False: tFecha.Text = ""
    cMoneda.BackColor = Colores.Gris: cMoneda.Enabled = False: cMoneda.Text = ""
    tCosto.BackColor = Colores.Gris: tCosto.Enabled = False: tCosto.Text = ""
    cIncoterm.BackColor = Colores.Gris: cIncoterm.Enabled = False: cIncoterm.Text = ""
    cITexto.BackColor = Colores.Gris: cITexto.Enabled = False: cITexto.Text = ""
    tComentario.BackColor = Colores.Gris: tComentario.Enabled = False: tComentario.Text = ""
    
    vsGrilla.Enabled = True: vsGrilla.BackColor = Colores.Blanco
    
End Sub

Private Sub HabilitoCampos()
    
    tNombre.BackColor = Inactivo: tNombre.Enabled = False
    tFecha.BackColor = Colores.Obligatorio: tFecha.Enabled = True
    cMoneda.BackColor = Colores.Obligatorio: cMoneda.Enabled = True
    tCosto.BackColor = Colores.Obligatorio: tCosto.Enabled = True
    cIncoterm.BackColor = Colores.Blanco: cIncoterm.Enabled = True
    cITexto.BackColor = Colores.Blanco: cITexto.Enabled = True
    tComentario.BackColor = Colores.Blanco: tComentario.Enabled = True
    
    vsGrilla.Enabled = False: vsGrilla.BackColor = Colores.Gris
    
End Sub

Private Sub CargoGrilla()
On Error GoTo ErrCI
Dim Cod As Integer, aValor As Long

    Screen.MousePointer = 11
    LimpioGrilla
    OcultoCampos
    
    Cons = " Select * From PrecioDeCosto " _
                    & " Left Outer Join Moneda On MonCodigo = PCoMoneda " _
                    & " Left Outer Join Incoterm On IncCodigo = PCoIncoterm " _
            & " Where PCoArticulo = " & Val(tNombre.Tag) _
            & " Order by PCoFecha Desc"
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not RsAux.EOF
        With vsGrilla
            .AddItem ""
            .Cell(flexcpText, vsGrilla.Rows - 1, 0) = Format(RsAux!PCoFecha, "dd/mm/yyyy hh:mm:ss")
            
            aValor = 0
            If Not IsNull(RsAux!PCoMoneda) Then
                .Cell(flexcpText, vsGrilla.Rows - 1, 1) = Trim(RsAux!MonSigno)
                aValor = RsAux!PCoMoneda
            End If
            .Cell(flexcpData, vsGrilla.Rows - 1, 1) = aValor
            
            .Cell(flexcpText, vsGrilla.Rows - 1, 2) = Format(RsAux!PCoImporte, FormatoMonedaP)
            
            aValor = 0
            If Not IsNull(RsAux!PCoIncoterm) Then
                .Cell(flexcpText, vsGrilla.Rows - 1, 3) = Trim(RsAux!IncNombre)
                aValor = RsAux!PCoIncoterm
            End If
            .Cell(flexcpData, vsGrilla.Rows - 1, 3) = aValor
            
            If Not IsNull(RsAux!PCoITexto) Then
                .Cell(flexcpText, vsGrilla.Rows - 1, 3) = Trim(.Cell(flexcpText, vsGrilla.Rows - 1, 3) & " " & Trim(RsAux!PCoITexto))
                .Cell(flexcpText, vsGrilla.Rows - 1, 5) = Trim(RsAux!PCoITexto)
            End If
            
            If Not IsNull(RsAux!PCoComentario) Then .Cell(flexcpText, vsGrilla.Rows - 1, 4) = Trim(RsAux!PCoComentario)
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
ErrCI:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos en la grilla.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Sub LimpioGrilla()
    
    With vsGrilla
        
        .Redraw = False
        .ExtendLastCol = True
        .Clear
        .Rows = 1: .Cols = 1
        .FormatString = "Fecha|>Moneda|>Costo|Incoterm|Cometarios|PCoITexto"
        .ColWidth(0) = 1400: .ColWidth(2) = 1000: .ColWidth(3) = 1600
        .AllowUserResizing = flexResizeColumns
        .ColHidden(5) = True
        .Redraw = True
        
    End With
    
End Sub

Private Sub vsgrilla_Click()
    If vsGrilla.MouseRow = 0 Then
        vsGrilla.ColSel = vsGrilla.MouseCol
        If vsGrilla.ColSort(vsGrilla.MouseCol) = flexSortGenericAscending Then
            vsGrilla.ColSort(vsGrilla.MouseCol) = flexSortGenericDescending
        Else
            vsGrilla.ColSort(vsGrilla.MouseCol) = flexSortGenericAscending
        End If
        vsGrilla.Sort = flexSortUseColSort
    End If
End Sub

Private Sub CargoCamposBD()
    
    RsAux!PCoArticulo = Val(tNombre.Tag)
    RsAux!PCoFecha = Format(tFecha.Text, sqlFormatoF) & " " & Time
    RsAux!PCoMoneda = cMoneda.ItemData(cMoneda.ListIndex)
    RsAux!PCoImporte = CCur(tCosto.Text)
    If cIncoterm.ListIndex <> -1 Then RsAux!PCoIncoterm = cIncoterm.ItemData(cIncoterm.ListIndex) Else RsAux!PCoIncoterm = Null
    If Trim(tComentario.Text) <> "" Then RsAux!PCoComentario = Trim(tComentario.Text) Else RsAux!PCoComentario = Null
    
End Sub

Private Sub BuscoArticuloPorCodigo(Articulo As Long)
On Error GoTo ErrBAPC
    Screen.MousePointer = 11
    Cons = "Select ArtID, ArtNombre From Articulo Where ArtCodigo = " & CLng(Articulo)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If RsAux.EOF Then
        Screen.MousePointer = 0
        MsgBox "No existe un artículo con ese código, o el mismo fue eliminado.", vbInformation, "ATENCIÓN"
        LimpioGrilla
        Botones False, False, False, False, False, Toolbar1, Me
        RsAux.Close
    Else
        tNombre.Text = Trim(RsAux!ArtNombre): tNombre.Tag = RsAux!ArtID
        RsAux.Close
        CargoGrilla
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrBAPC:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el artículo por código."
    Screen.MousePointer = 0
End Sub
Private Sub BuscoArticuloPorNombre()

    Cons = "Select ArtCodigo as 'Código', ArtNombre as 'Nombre' From Articulo Where ArtNombre Like '" & Replace(tNombre.Text, " ", "%") & "%'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        Screen.MousePointer = 0
        MsgBox "No existe un artículo con ese nombre.", vbInformation, "ATENCIÓN"
    Else
        RsAux.MoveNext
        If RsAux.EOF Then
            RsAux.MoveFirst
            tNombre.Tag = RsAux(0)
            RsAux.Close
        Else
            RsAux.Close
            Dim objLista As New clsListadeAyuda
            If objLista.ActivarAyuda(cBase, Cons, 5000, 0, "Lista de Artículos") > 0 Then
                tNombre.Tag = objLista.RetornoDatoSeleccionado(0)
            End If
            Set objLista = Nothing
        End If
        Screen.MousePointer = 0
        If Val(tNombre.Tag) > 0 Then BuscoArticuloPorCodigo Val(tNombre.Tag)
    End If

End Sub
