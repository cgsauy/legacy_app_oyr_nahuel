VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form frmTrasladoEspecial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferencia de Mercadería Especial"
   ClientHeight    =   5205
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6630
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTrasladoEspecial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir del formulario"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "imprimir"
            Object.ToolTipText     =   "Imprimir entrega"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "entregar"
            Object.ToolTipText     =   "Entregar Mercadería"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox tMinuto 
      Height          =   285
      Left            =   4320
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   4740
      Width           =   435
   End
   Begin VB.TextBox tHora 
      Height          =   285
      Left            =   3660
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   4740
      Width           =   435
   End
   Begin VB.TextBox tFecha 
      Height          =   285
      Left            =   1860
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   4740
      Width           =   1215
   End
   Begin VB.TextBox tUsuario 
      Height          =   285
      Left            =   5640
      TabIndex        =   21
      Top             =   4740
      Width           =   915
   End
   Begin VB.TextBox tComentario 
      Height          =   285
      Left            =   1140
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   4380
      Width           =   5415
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   2115
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   3731
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
      SelectionMode   =   0
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
   Begin AACombo99.AACombo cEstado 
      Height          =   315
      Left            =   4320
      TabIndex        =   10
      Top             =   1500
      Width           =   2235
      _ExtentX        =   3942
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
   Begin VB.TextBox tCantidad 
      Height          =   285
      Left            =   1140
      MaxLength       =   8
      TabIndex        =   8
      Top             =   1500
      Width           =   1275
   End
   Begin VB.TextBox tArticulo 
      Height          =   315
      Left            =   1140
      TabIndex        =   6
      Top             =   1080
      Width           =   5415
   End
   Begin AACombo99.AACombo cLDestino 
      Height          =   315
      Left            =   4320
      TabIndex        =   3
      Top             =   660
      Width           =   2235
      _ExtentX        =   3942
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
   Begin AACombo99.AACombo cLOrigen 
      Height          =   315
      Left            =   1140
      TabIndex        =   1
      Top             =   660
      Width           =   2295
      _ExtentX        =   4048
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6240
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasladoEspecial.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasladoEspecial.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasladoEspecial.frx":0736
            Key             =   "Total"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasladoEspecial.frx":0A50
            Key             =   "Parcial"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasladoEspecial.frx":0D6A
            Key             =   "No"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasladoEspecial.frx":1084
            Key             =   "NoDoy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasladoEspecial.frx":139E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasladoEspecial.frx":14B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasladoEspecial.frx":15C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasladoEspecial.frx":16D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasladoEspecial.frx":17E6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   22
      Top             =   4740
      Width           =   195
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "&Hora:"
      Height          =   255
      Left            =   3180
      TabIndex        =   17
      Top             =   4740
      Width           =   555
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "&Fecha del Movimiento:"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   4740
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario:"
      Height          =   195
      Left            =   4980
      TabIndex        =   20
      Top             =   4740
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Co&mentario:"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   4380
      Width           =   975
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   " A&rtículos Ingresados"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   6375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Estado:"
      Height          =   195
      Left            =   3540
      TabIndex        =   9
      Top             =   1500
      Width           =   795
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Cantidad:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1500
      Width           =   915
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Artículo:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Destino:"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   660
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Local Origen:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   660
      Width           =   1035
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
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
      Begin VB.Menu MnuOpLinea2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
         Visible         =   0   'False
      End
      Begin VB.Menu MnuOpLinea3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "&Salir del formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmTrasladoEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cEstado_GotFocus()
On Error Resume Next
    With cEstado
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub cEstado_KeyPress(KeyAscii As Integer)
Dim iStock As Long

    If KeyAscii = vbKeyReturn And Val(tArticulo.Tag) > 0 And IsNumeric(tCantidad.Text) And cEstado.ListIndex > -1 Then
        iStock = StockLocalArticuloyEstado(tArticulo.Tag, cEstado.ItemData(cEstado.ListIndex))
        If iStock < CLng(tCantidad.Text) Then
            If MsgBox("No hay tantos artículos en stock. ¿Desea continuar de todas formas?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
        End If
        InsertoArticuloEnlaLista (iStock)
        LimpioIngreso
        tArticulo.SetFocus
    ElseIf KeyAscii = vbKeyReturn Then
        MsgBox "Los datos son incompletos o incorrectos.", vbExclamation, "ATENCIÓN"
    End If

End Sub

Private Sub cLDestino_GotFocus()
On Error Resume Next
    With cLDestino
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub cLDestino_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn And cLDestino.ListIndex > -1 Then tArticulo.SetFocus
End Sub

Private Sub cLOrigen_GotFocus()
On Error Resume Next
    With cLOrigen
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub cLOrigen_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn And cLOrigen.ListIndex > -1 Then cLDestino.SetFocus
End Sub

Private Sub Form_Activate()
On Error Resume Next
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    
    ObtengoSeteoForm Me
    EstadoObjetos False
    
    With vsConsulta
        .Redraw = False
        .Editable = False: .Rows = 1: .Cols = 1: .ExtendLastCol = True
        .FormatString = "Cantidad|<Estado|<Artículo|En Stock"
        .ColWidth(1) = 1000: .ColWidth(2) = 3500
        .ColHidden(3) = True
        .Redraw = True
    End With
    
    'Cargo los locales
    Cons = "Select LocCodigo, LocNombre From Local Order by LocNombre"
    CargoCombo Cons, cLOrigen
    CargoCombo Cons, cLDestino

    'Cargo Estados
    Cons = "Select EsMCodigo, EsMAbreviacion From EstadoMercaderia " _
        & " Where EsMBajaStockTotal = 0 Order by EsMAbreviacion"
    CargoCombo Cons, cEstado, ""
    Exit Sub
    
errLoad:
    MsgBox "Ocurrió el siguiente error al iniciar el formulario: " & vbCr & Trim(Err.Description), vbExclamation, "ATENCIÓN"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    GuardoSeteoForm Me
    CierroConexion
    Set miConexion = Nothing
End Sub

Private Sub Label1_Click()
On Error Resume Next
    With cLOrigen
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label2_Click()
On Error Resume Next
    With cLDestino
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label3_Click()
On Error Resume Next
    With tArticulo
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label4_Click()
On Error Resume Next
    With tCantidad
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label5_Click()
On Error Resume Next
    With cEstado
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label6_Click()
On Error Resume Next
    vsConsulta.SetFocus
End Sub

Private Sub Label7_Click()
On Error Resume Next
    With tComentario
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label8_Click()
On Error Resume Next
    With tUsuario
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuNuevo_Click()
    AccionNuevo
End Sub

Private Sub MnuSalir_Click()
    Unload Me
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = ""
End Sub

Private Sub tArticulo_GotFocus()
On Error Resume Next
    With tArticulo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tArticulo_KeyPress(KeyAscii As Integer)
On Error GoTo errTA
    If KeyAscii = vbKeyReturn Then
        If Trim(tArticulo.Text) <> "" And tArticulo.Tag = "" Then
            If IsNumeric(tArticulo.Text) Then
                BuscoArticuloPorCodigo Val(tArticulo.Text)
            Else
                BuscoArticuloPorNombre
            End If
            If Val(tArticulo.Tag) > 0 Then tCantidad.SetFocus
        Else
            If Trim(tArticulo.Text) = "" Then vsConsulta.SetFocus Else tCantidad.SetFocus
        End If
    End If
    Exit Sub
errTA:
    MsgBox "Ocurrió el siguiente error al procesar el artículo: " & vbCr & Trim(Err.Description), vbExclamation, "ATENCIÓN"
End Sub

Private Sub BuscoArticuloPorNombre()
On Error GoTo ErrBAN
Dim aCodigo As Long: aCodigo = 0

    Screen.MousePointer = vbHourglass
    Cons = "Select ArtCodigo as 'Código', ArtNombre As 'Nombre' From Articulo" _
        & " Where ArtNombre LIKE '" & Replace(tArticulo.Text, " ", "%") & "%'" _
        & " Order by ArtNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No existe un artículo para el dato ingresado.", vbInformation, "ATENCIÓN"
    Else
        RsAux.MoveNext
        If RsAux.EOF Then
            RsAux.MoveFirst
            aCodigo = RsAux(0)
        Else
            Dim objLista As New clsListadeAyuda
            If objLista.ActivarAyuda(cBase, Cons, 4500, 0, "Artículos") > 0 Then
                aCodigo = objLista.RetornoDatoSeleccionado(0)
            End If
            Set objLista = Nothing
        End If
        RsAux.Close
    End If
    If aCodigo > 0 Then BuscoArticuloPorCodigo aCodigo Else tArticulo.Tag = ""
    Screen.MousePointer = 0
    Exit Sub
    
ErrBAN:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub BuscoArticuloPorCodigo(aCodigo As Long)
On Error GoTo ErrBAC
    Screen.MousePointer = 11
    tArticulo.Text = ""
    Cons = "Select * From Articulo Where ArtCodigo = " & aCodigo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsAux.EOF Then
        tArticulo.Text = "(" & Format(RsAux!ArtCodigo, "000,000") & ") " & Trim(RsAux!ArtNombre)
        tArticulo.Tag = RsAux!ArtID
    End If
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub
ErrBAC:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el artículo por código.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tCantidad_GotFocus()
On Error Resume Next
    With tCantidad
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tCantidad_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tCantidad.Text) Then
            If CLng(tCantidad.Text) > 0 Then
                cEstado.SetFocus
                If cEstado.ListIndex = -1 Then BuscoCodigoEnCombo cEstado, CLng(paEstadoArticuloEntrega)
            Else
                MsgBox "Debe ingresar un número mayor que cero.", vbExclamation, "ATENCIÓN"
            End If
        ElseIf Trim(tCantidad.Text) <> "" Then
            MsgBox "El formato no es numérico.", vbExclamation, "ATENCIÓN"
        End If
    End If
End Sub

Private Sub tCantidad_LostFocus()
On Error Resume Next
    If IsNumeric(tCantidad.Text) Then
        If CLng(tCantidad.Text) < 0 Then tCantidad.Text = ""
    End If
End Sub

Private Function StockLocalArticuloyEstado(lnArticulo As Long, iEstado As Integer) As Integer
On Error GoTo errSTL
Dim RS As rdoResultset

    Screen.MousePointer = vbHourglass
    StockLocalArticuloyEstado = 0

    Cons = "Select Sum(StLCantidad) From StockLocal " _
        & " Where StLArticulo = " & lnArticulo _
        & " And StLLocal = " & cLOrigen.ItemData(cLOrigen.ListIndex) & " And StLEstado = " & iEstado
        
    Set RS = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not RS.EOF Then If Not IsNull(RS(0)) Then StockLocalArticuloyEstado = RS(0)
    RS.Close
    Screen.MousePointer = vbDefault
    Exit Function
        
errSTL:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error inesperado al buscar el stock del local."

End Function


Private Sub InsertoArticuloEnlaLista(Stock As Integer)
    
    On Error GoTo ErrClave
    
    With vsConsulta
        'Veo si ya inserte el artículo para el estado seleccionado.
        For I = 1 To .Rows - 1
            If CLng(.Cell(flexcpData, I, 0)) = Val(tArticulo.Tag) _
                And CInt(.Cell(flexcpData, I, 1)) = cEstado.ItemData(cEstado.ListIndex) Then
                MsgBox "Ya se inserto ese artículo con el estado seleccionado, verifique.", vbExclamation, "ATENCIÓN": Exit Sub
            End If
        Next I
    
        .AddItem tCantidad.Text
        .Cell(flexcpText, .Rows - 1, 1) = cEstado.Text
        .Cell(flexcpText, .Rows - 1, 2) = Trim(tArticulo.Text)
        .Cell(flexcpText, .Rows - 1, 3) = Stock
        'Data
        .Cell(flexcpData, .Rows - 1, 0) = Val(tArticulo.Tag)
        .Cell(flexcpData, .Rows - 1, 1) = cEstado.ItemData(cEstado.ListIndex)
    End With
    Exit Sub

ErrIAEL:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error inesperado al ingresar el artículo en la lista."
    Exit Sub
    
ErrClave:
    Screen.MousePointer = vbDefault
    MsgBox "El artículo ya fue ingresado con ese estado, verífique.", vbCritical, "ATENCIÓN"
    
End Sub

Private Sub LimpioIngreso()
    tArticulo.Text = "": tArticulo.Tag = "": tCantidad.Text = "": cEstado.Text = ""
End Sub

Private Sub EstadoObjetos(sActivar As Boolean)
On Error Resume Next
Dim auxColor As Long

    cLOrigen.Enabled = sActivar: cLDestino.Enabled = sActivar: tArticulo.Enabled = sActivar
    tCantidad.Enabled = sActivar: cEstado.Enabled = sActivar: vsConsulta.Enabled = sActivar
    tComentario.Enabled = sActivar: tUsuario.Enabled = sActivar: tFecha.Enabled = sActivar
    tHora.Enabled = sActivar: tMinuto.Enabled = sActivar
    
    cLOrigen.Text = "": cLDestino.Text = "": tArticulo.Text = ""
    tCantidad.Text = "": cEstado.Text = "": vsConsulta.Rows = 1
    tComentario.Text = "": tUsuario.Text = "": tFecha.Text = "": tHora.Text = "": tMinuto.Text = ""
    
    If sActivar Then auxColor = vbWindowBackground Else auxColor = vbButtonFace
    
    cLOrigen.BackColor = auxColor: cLDestino.BackColor = auxColor: tArticulo.BackColor = auxColor
    tCantidad.BackColor = auxColor: cEstado.BackColor = auxColor: vsConsulta.BackColor = auxColor
    tComentario.BackColor = auxColor: tUsuario.BackColor = auxColor
    tFecha.BackColor = auxColor: tHora.BackColor = auxColor: tMinuto.BackColor = auxColor
    
    MnuNuevo.Enabled = Not sActivar: Toolbar1.Buttons("nuevo").Enabled = Not sActivar
    MnuGrabar.Enabled = sActivar: Toolbar1.Buttons("grabar").Enabled = sActivar
    MnuCancelar.Enabled = sActivar: Toolbar1.Buttons("cancelar").Enabled = sActivar
    
End Sub

Private Sub tComentario_GotFocus()
On Error Resume Next
    With tComentario
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tFecha.SetFocus
End Sub


Private Sub tFecha_GotFocus()
On Error Resume Next
    With tFecha
        If Not IsDate(.Text) Then .Text = Format(Date, "dd/mm/yyyy")
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tHora.SetFocus
End Sub

Private Sub tFecha_LostFocus()
On Error Resume Next
    If Trim(tFecha.Text) <> "" And Not IsDate(tFecha.Text) Then MsgBox "No ingreso una fecha válida.", vbInformation, "ATENCIÓN" Else tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy")
End Sub

Private Sub tHora_GotFocus()
On Error Resume Next
    With tHora
        If .Text = "" Then .Text = Format(Time, "hh")
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tHora_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Not IsNumeric(tHora.Text) And IsDate(tFecha.Text) Then
            MsgBox "Ingrese la hora del movimiento.", vbExclamation, "ATENCIÓN"
        Else
            tMinuto.SetFocus
        End If
    End If
End Sub

Private Sub tHora_LostFocus()
On Error Resume Next
    If Val(tHora.Text) > 23 Or Val(tHora.Text) < 0 Then MsgBox "La hora ingresada no es correcta.", vbExclamation, "ATENCIÓN": tHora.SetFocus
End Sub
Private Sub tMinuto_GotFocus()
On Error Resume Next
    With tMinuto
        If .Text = "" Then .Text = Mid(Time, 4, 2)
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tMinuto_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If Not IsNumeric(tMinuto.Text) And IsDate(tFecha.Text) Then
            MsgBox "Ingrese los minutos del movimiento.", vbExclamation, "ATENCIÓN"
        Else
            tUsuario.SetFocus
        End If
    End If
End Sub

Private Sub tMinuto_LostFocus()
On Error Resume Next
    If Val(tMinuto.Text) > 59 Or Val(tMinuto.Text) < 0 Then MsgBox "Los minutos ingresados no son correctos.", vbExclamation, "ATENCIÓN": tMinuto.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
On Error Resume Next
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "grabar": AccionGrabar
        Case "salir": Unload Me
        Case "cancelar": AccionCancelar
    End Select
End Sub

Private Sub AccionCancelar()
    EstadoObjetos False
End Sub

Private Sub AccionGrabar()
    If MsgBox("¿Confirma almacenar la información ingresada.", vbQuestion + vbYesNo, "Grabar") = vbYes Then
        'Válido los datos
        If cLOrigen.ListIndex = -1 Then MsgBox "El origen no es correcto.", vbInformation, "ATENCIÓN": cLOrigen.SetFocus: Exit Sub
        If cLDestino.ListIndex = -1 Then MsgBox "El destino no es correcto.", vbInformation, "ATENCIÓN": cLDestino.SetFocus: Exit Sub
        If cLOrigen.ItemData(cLOrigen.ListIndex) = cLDestino.ItemData(cLDestino.ListIndex) Then MsgBox "Ingreso el mismo local.", vbInformation, "ATENCIÓN": cLOrigen.SetFocus: Exit Sub
        If vsConsulta.Rows = 1 Then MsgBox "No hay artículos ingresados.", vbInformation, "ATENCIÓN": tArticulo.SetFocus: Exit Sub
        If Not IsDate(tFecha.Text) Then
            If MsgBox("No ingreso una fecha para el movimiento, se almacenaran con la fecha del sistema." & vbCr & "¿Desea continuar de todas formas?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
        Else
            If Not IsNumeric(tHora.Text) Then MsgBox "Ingrese la hora del movimiento.", vbInformation, "ATENCIÓN": tHora.SetFocus: Exit Sub
            If Not IsNumeric(tMinuto.Text) Then MsgBox "Ingrese los minutos del movimiento.", vbInformation, "ATENCIÓN": tMinuto.SetFocus: Exit Sub
            If Val(tHora.Text) < 0 Or Val(tHora.Text) > 23 Then MsgBox "La hora ingresada no es válida", vbInformation, "ATENCIÓN": tHora.SetFocus: Exit Sub
            If Val(tMinuto.Text) < 0 Or Val(tMinuto.Text) > 59 Then MsgBox "Los mintuos ingresados no son validos", vbInformation, "ATENCIÓN": tMinuto.SetFocus: Exit Sub
        End If
        If Val(tUsuario.Tag) <= 0 Then MsgBox "No ingreso un usuario.", vbInformation, "ATENCIÓN": tUsuario.SetFocus: Exit Sub
        GraboTraslado
    End If
End Sub

Private Sub AccionNuevo()
On Error Resume Next
    EstadoObjetos True
    tArticulo.SetFocus
    cLOrigen.SetFocus
End Sub

Private Sub tUsuario_GotFocus()
On Error Resume Next
    With tUsuario
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)
On Error Resume Next
    tUsuario.Tag = vbNullString
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tUsuario.Text) Then
            tUsuario.Tag = BuscoUsuarioDigito(CInt(tUsuario.Text), True)
            If CInt(tUsuario.Tag) > 0 Then AccionGrabar Else tUsuario.Tag = vbNullString
        Else
            MsgBox "El formato del código no es numérico.", vbExclamation, "ATENCIÓN"
            tUsuario.SetFocus
        End If
    End If
End Sub

Private Sub GraboTraslado()
Dim strFecha As String
Dim iAux As Long
Dim idTipoOrigen As Integer, idTipoDestino As Integer

    Screen.MousePointer = 11
    FechaDelServidor
    
    idTipoOrigen = 0: idTipoDestino = 0
    
    Cons = "Select * From Local Where LocCodigo IN (" & cLOrigen.ItemData(cLOrigen.ListIndex) & ", " & cLDestino.ItemData(cLDestino.ListIndex) & ")"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case RsAux!LocCodigo
            Case cLOrigen.ItemData(cLOrigen.ListIndex): idTipoOrigen = RsAux!LocTipo
            Case cLDestino.ItemData(cLDestino.ListIndex): idTipoDestino = RsAux!LocTipo
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    On Error GoTo ErrBT
    cBase.BeginTrans
    On Error GoTo ErrRelajo
    
    'Primero inserto en la tabla traspaso, luego obtengo su código.
    'Seguido inserto los renglones y cambio los movimientos de stock.
    
    Cons = "Select * From Traspaso Where TraCodigo = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    RsAux.AddNew
    If tFecha.Text <> "" Then
        strFecha = tFecha.Text & " " & tHora.Text & ":" & tMinuto.Text & ":00"
    Else
        strFecha = gFechaServidor
    End If
    RsAux!TraFecha = Format(strFecha, sqlFormatoFH)
    RsAux!TraLocalOrigen = cLOrigen.ItemData(cLOrigen.ListIndex)
    RsAux!TraLocalDestino = cLDestino.ItemData(cLDestino.ListIndex)
    If Trim(tComentario.Text) <> "" Then RsAux!TraComentario = tComentario.Text
    RsAux!TraFechaEntregado = Format(strFecha, sqlFormatoFH)
    RsAux!TraUsuarioReceptor = CInt(tUsuario.Tag)
    RsAux!TraUsuarioFinal = CInt(tUsuario.Tag)
    RsAux!TraUsuarioInicial = CInt(tUsuario.Tag)
    RsAux!TraFImpreso = Format(strFecha, sqlFormatoFH)
    RsAux!TraTerminal = paCodigoDeTerminal
    RsAux.Update
    RsAux.Close
    
    'Saco el código del insertado.
    Cons = "Select MAX(TraCodigo) From Traspaso Where TraCodigo > " & iAux _
        & " And TraLocalOrigen = " & cLOrigen.ItemData(cLOrigen.ListIndex) _
        & " And TraLocalDestino = " & cLDestino.ItemData(cLDestino.ListIndex) _
        & " And TraUsuarioFinal = " & Val(tUsuario.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not IsNull(RsAux(0)) Then iAux = RsAux(0) Else iAux = 0
    RsAux.Close
    
    'A la fecha del servidor le pongo la fecha que ingreso.
    gFechaServidor = strFecha
    
    'Inserto los artículos en la tabla renglontraspaso y hago los mov. físicos.
    With vsConsulta
        For I = 1 To .Rows - 1
            
            If CInt(.Cell(flexcpText, I, 0)) > 0 Then
                
                'Como pendiente pongo cero.
                Cons = "Insert into RenglonTraspaso (RTrTraspaso, RTrArticulo, RTrEstado, RTrCantidad, RTrPendiente)" _
                    & " Values (" & iAux & ", " & Val(.Cell(flexcpData, I, 0)) & ", " & Val(.Cell(flexcpData, I, 1)) & ", " & Val(.Cell(flexcpText, I, 0)) & ", 0)"
                cBase.Execute (Cons)
                
                'Dar la baja al origen
                Cons = "Select * From StockLocal " _
                    & " Where StLArticulo = " & Val(.Cell(flexcpData, I, 0)) _
                    & " And StlTipoLocal = " & idTipoOrigen & " And StLLocal = " & cLOrigen.ItemData(cLOrigen.ListIndex) _
                    & " And StLEstado = " & Val(.Cell(flexcpData, I, 1))
                
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
                If RsAux.EOF Then
                    RsAux.AddNew
                    RsAux!StLArticulo = Val(.Cell(flexcpData, I, 0))
                    RsAux!StLTipoLocal = idTipoOrigen
                    RsAux!StlLocal = cLOrigen.ItemData(cLOrigen.ListIndex)
                    RsAux!StLEstado = Val(.Cell(flexcpData, I, 1))
                    RsAux!StLCantidad = Val(.Cell(flexcpText, I, 0)) * -1
                    RsAux.Update
                Else
                    If RsAux!StLCantidad - CInt(.Cell(flexcpText, I, 0)) = 0 Then
                        RsAux.Delete
                    Else
                        RsAux.Edit
                        RsAux!StLCantidad = RsAux!StLCantidad - CInt(.Cell(flexcpText, I, 0))
                        RsAux.Update
                    End If
                End If
                RsAux.Close
                MarcoMovimientoStockFisico CLng(tUsuario.Tag), idTipoOrigen, cLOrigen.ItemData(cLOrigen.ListIndex), Val(.Cell(flexcpData, I, 0)), CInt(.Cell(flexcpText, I, 0)), Val(.Cell(flexcpData, I, 1)), -1, TipoDocumento.Traslados, iAux
                
                'Inserto en el destino.
                'Hago recepción.
                Cons = "Select * From StockLocal " _
                    & " Where StLArticulo = " & Val(.Cell(flexcpData, I, 0)) _
                    & " And StlTipoLocal = " & idTipoDestino & " And StLLocal = " & cLDestino.ItemData(cLDestino.ListIndex) _
                    & " And StLEstado = " & Val(.Cell(flexcpData, I, 1))
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If RsAux.EOF Then
                    RsAux.AddNew
                    RsAux!StLArticulo = Val(.Cell(flexcpData, I, 0))
                    RsAux!StLTipoLocal = idTipoDestino
                    RsAux!StlLocal = cLDestino.ItemData(cLDestino.ListIndex)
                    RsAux!StLEstado = Val(.Cell(flexcpData, I, 1))
                    RsAux!StLCantidad = Val(.Cell(flexcpText, I, 0))
                    RsAux.Update
                Else
                    RsAux.Edit
                    RsAux!StLCantidad = RsAux!StLCantidad + Val(.Cell(flexcpText, I, 0))
                    RsAux.Update
                End If
                RsAux.Close
                MarcoMovimientoStockFisico CLng(tUsuario.Tag), idTipoDestino, cLDestino.ItemData(cLDestino.ListIndex), Val(.Cell(flexcpData, I, 0)), Val(.Cell(flexcpText, I, 0)), Val(.Cell(flexcpData, I, 1)), 1, TipoDocumento.Traslados, iAux
            End If
        Next
    End With
    cBase.CommitTrans
    
    AccionCancelar
    Screen.MousePointer = 0
    Exit Sub
    
ErrBT:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al intentar iniciar la transaccion.", Trim(Err.Description)
    Exit Sub
    
ErrRelajo:
    Resume Resumo
    
Resumo:
    cBase.RollbackTrans
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al intentar grabar los datos.", Err.Description
    
End Sub

Private Sub vsConsulta_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    With vsConsulta
        If vsConsulta.Rows > 1 Then
            Select Case KeyCode
                
                Case vbKeyAdd: .Cell(flexcpText, .Row, 0) = CLng(.Cell(flexcpText, .Row, 0)) + 1
                Case vbKeySubtract
                    If CLng(.Cell(flexcpText, .Row, 0)) > 1 Then .Cell(flexcpText, .Row, 0) = CLng(.Cell(flexcpText, .Row, 0)) - 1
                Case vbKeyDelete: .RemoveItem .Row
                Case vbKeyReturn: tComentario.SetFocus
                   
            End Select
            
        End If
    End With

End Sub

