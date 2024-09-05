VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmLinea 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Líneas"
   ClientHeight    =   4320
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMaLinea.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   3570
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar tooMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
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
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   2
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
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox tNombre 
      Height          =   285
      Left            =   900
      MaxLength       =   20
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   2535
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsDato 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   5318
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
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
   Begin ComctlLib.StatusBar staMsg 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   4065
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Nombre:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   675
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4860
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaLinea.frx":164A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaLinea.frx":175C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaLinea.frx":186E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaLinea.frx":1980
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaLinea.frx":1A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaLinea.frx":1DAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaLinea.frx":20C6
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
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuOpLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGrabar 
         Caption         =   "&Grabar"
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "&Cancelar"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuSalForm 
         Caption         =   "Del formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmLinea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bEdicion As Boolean

Private Sub Form_Load()
On Error GoTo errLoad
    
    ObtengoSeteoForm Me, 500, 500
    Botones True, False, False, False, False, tooMenu, Me
    LimpioDatos
    bEdicion = False
    With vsDato
        .Rows = 1: .Cols = 1: .ExtendLastCol = True
        .FormatString = "<Nombre"
        .ColWidth(0) = 1500
    End With
    Screen.MousePointer = 0
    Exit Sub
errLoad:
    clsGeneral.OcurrioError "Ocurrió un error al iniciar el formulario.", Trim(Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    CierroConexion
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    GuardoSeteoForm Me
End Sub

Private Sub Label1_Click()
On Error Resume Next
    With tNombre
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
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


Private Sub MnuSalForm_Click()
    Unload Me
End Sub

Private Sub tNombre_GotFocus()
On Error Resume Next
    With tNombre
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If bEdicion Then
            AccionGrabar
        Else
            AccionBuscar
        End If
    End If
End Sub

Private Sub tooMenu_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "salir": Unload Me
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
    End Select
End Sub

Private Sub AccionNuevo()
    vsDato.Enabled = False
    LimpioDatos
    Botones False, False, False, True, True, tooMenu, Me
    bEdicion = True
    tNombre.SetFocus
End Sub

Private Sub AccionModificar()
    LimpioDatos
    Cons = "Select * From CodigoTexto Where Codigo = " & vsDato.Cell(flexcpData, vsDato.Row, 0)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "Otra terminal pudo eliminar la línea seleccionada.", vbExclamation, "ATENCIÓN"
        Exit Sub
    Else
        tNombre.Text = Trim(RsAux!Texto)
        tNombre.Tag = RsAux!Codigo
        RsAux.Close
    End If
    Botones False, False, False, True, True, tooMenu, Me
    vsDato.Enabled = False
    bEdicion = True
    tNombre.SetFocus
    
End Sub

Private Sub AccionEliminar()

    If MsgBox("¿Confirma eliminar la línea " & Trim(vsDato.Text) & " ?", vbQuestion + vbYesNo, "Eliminar") = vbYes Then
        Screen.MousePointer = 11
        Cons = "Select * From FleteEmbarque Where FEmLinea = " & vsDato.Cell(flexcpData, vsDato.Row, 0)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            RsAux.Close
            MsgBox "Esta línea esta relacionad en flete de embarque, no podrá eliminarla mientras exista la relación.", vbInformation, "ATENCIÓN"
            Screen.MousePointer = 0
            Exit Sub
        End If
        RsAux.Close
        On Error GoTo errBT
        cBase.CommitTrans
        On Error GoTo errRB
        Cons = "Select * From CodigoTexto Where Codigo = " & vsDato.Cell(flexcpData, vsDato.Row, 0)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.Delete
        RsAux.Close
        cBase.CommitTrans
        vsDato.Rows = 1
        Screen.MousePointer = 0
        AccionCancelar
    End If
    Exit Sub

errBT:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar la transacción.", Trim(Err.Description)
    Screen.MousePointer = 0
    Exit Sub

RollB:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrió un error al almacenar la información.", Trim(Err.Description)
    Screen.MousePointer = 0

errRB:
    Resume RollB

End Sub

Private Sub AccionGrabar()
    If MsgBox("¿Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
        If ValidoDatos Then GraboEnBD
    End If
End Sub

Private Sub AccionCancelar()
On Error Resume Next
    
    LimpioDatos
    vsDato.Enabled = True
    vsDato.SetFocus
    bEdicion = False
    If vsDato.Rows > 1 Then
        Botones True, True, True, False, False, tooMenu, Me
    Else
        Botones True, False, False, False, False, tooMenu, Me
    End If
    
End Sub

Private Sub GraboEnBD()
Dim lpos As Long, lCol As Long
    
    Screen.MousePointer = 11
    On Error GoTo errBT
    '-----------------------------------------------------
    cBase.BeginTrans
    On Error GoTo errRB
    Cons = "Select * From CodigoTexto Where Codigo = " & Val(tNombre.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.AddNew
    Else
        RsAux.Edit
    End If
    RsAux!Tipo = 67
    RsAux!Texto = Trim(tNombre.Text)
    RsAux.Update
    RsAux.Close
    cBase.CommitTrans
    '-----------------------------------------------------
    AccionBuscar
    AccionCancelar
    Screen.MousePointer = 0
    Exit Sub
errBT:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar la transacción.", Trim(Err.Description)
    Screen.MousePointer = 0
    Exit Sub


RollB:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrió un error al almacenar la información.", Trim(Err.Description)
    Screen.MousePointer = 0
    Exit Sub

errRB:
    Resume RollB
    
End Sub

Private Sub LimpioDatos()
    tNombre.Text = ""
    tNombre.Tag = ""
End Sub

Private Sub AccionBuscar(Optional Flete As Long = 0)
On Error GoTo errAB
Dim lValor As Long
    
    Screen.MousePointer = 11
    vsDato.Rows = 1
    Botones True, False, False, False, False, tooMenu, Me
    Cons = "Select * From CodigoTexto Where Tipo = 67 And Texto Like '" & clsGeneral.Replace(tNombre.Text, " ", "%") & "%'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        With vsDato
            .AddItem Trim(RsAux!Texto)
            lValor = RsAux!Codigo: .Cell(flexcpData, .Rows - 1, 0) = lValor
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    If vsDato.Rows > 1 Then
        tNombre.Text = "": tNombre.Tag = ""
        vsDato.Select 1, 0
        Botones True, True, True, False, False, tooMenu, Me
    Else
        MsgBox "No hay líneas ingresadas con esas características.", vbInformation, "ATENCIÓN"
    End If
    Screen.MousePointer = 0
    Exit Sub
errAB:
    clsGeneral.OcurrioError "Ocurrió un error al buscar.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function ValidoDatos() As Boolean
    ValidoDatos = False
    If Trim(tNombre.Text) = "" Then
        MsgBox "El nombre es obligatorio.", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    ValidoDatos = True
End Function

Private Sub vsDato_SelChange()
    If vsDato.Row >= 1 Then
        Botones True, True, True, False, False, tooMenu, Me
    Else
        Botones True, False, False, False, False, tooMenu, Me
    End If
End Sub
