VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form CamEstadoMercaderia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Estado de Mercadería"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCamEstadoMercaderia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   5205
   StartUpPosition =   1  'CenterOwner
   Begin AACombo99.AACombo cLocal 
      Height          =   315
      Left            =   3000
      TabIndex        =   3
      Top             =   600
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   556
      BackColor       =   12648447
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
   Begin AACombo99.AACombo cEstadoViejo 
      Height          =   315
      Left            =   3060
      TabIndex        =   9
      Top             =   1320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      BackColor       =   12648447
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
   Begin AACombo99.AACombo cEstadoNuevo 
      Height          =   315
      Left            =   1020
      TabIndex        =   11
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BackColor       =   12648447
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   3000
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stock de Camiones"
      Height          =   975
      Left            =   120
      TabIndex        =   27
      Top             =   2160
      Width           =   4935
      Begin VB.CommandButton bAyuda 
         Appearance      =   0  'Flat
         Caption         =   "A&yuda"
         Height          =   375
         Left            =   4080
         Picture         =   "frmCamEstadoMercaderia.frx":0442
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox tCodigoTraspaso 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3120
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox tCodigoImpresion 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3120
         TabIndex        =   17
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton opOpciones 
         Caption         =   "En&víos"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton opOpciones 
         Caption         =   "&Traslados"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Códi&go de Traslado:"
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Código de &Impresión:"
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   600
         Width           =   1695
      End
   End
   Begin MSComctlLib.ListView lvArticulo 
      Height          =   1095
      Left            =   120
      TabIndex        =   20
      Top             =   3480
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1931
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cant."
         Object.Width           =   1050
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Artículo"
         Object.Width           =   2981
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Estado viejo"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nuevo"
         Object.Width           =   1252
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tipo"
         Object.Width           =   952
      EndProperty
   End
   Begin VB.TextBox tUsuario 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   24
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox tComentario 
      Height          =   285
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   22
      Top             =   4680
      Width           =   3975
   End
   Begin VB.TextBox tCantidad 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1020
      MaxLength       =   5
      TabIndex        =   7
      Top             =   1320
      Width           =   735
   End
   Begin VB.ComboBox cArticulo 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1020
      Style           =   1  'Simple Combo
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   960
      Width           =   3615
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   5415
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.TextBox tFecha 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1020
      MaxLength       =   12
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCamEstadoMercaderia.frx":0884
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCamEstadoMercaderia.frx":0996
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCamEstadoMercaderia.frx":0AA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCamEstadoMercaderia.frx":0BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCamEstadoMercaderia.frx":0CCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCamEstadoMercaderia.frx":0DDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCamEstadoMercaderia.frx":10F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCamEstadoMercaderia.frx":120A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " A&rtículos ingresados"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   4935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Come&ntario:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Ca&mbiar a:"
      Height          =   255
      Left            =   180
      TabIndex        =   10
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Estado:"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Cantidad:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Artículo:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Fecha:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Local:"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   600
      Width           =   615
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
         Visible         =   0   'False
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Shortcut        =   ^E
         Visible         =   0   'False
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
         Caption         =   "&Volver al Formulario Anterior"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "CamEstadoMercaderia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TipoLocal
    Camion = 1
    Deposito = 2
End Enum

Private Enum TipoControlMercaderia
    CambioEstado = 1
    EntregaMercaderia = 2
End Enum


Private Const FormatoFH = "mm/dd/yyyy hh:mm:ss"
Dim itmX As ListItem
Dim Msg As String

Private Sub bAyuda_Click()
On Error GoTo ErrBA

    If Trim(tCantidad.Text) = vbNullString Then
        MsgBox " Ingrese la cantidad a modificar.", vbExclamation, "ATENCIÓN"
        tCantidad.SetFocus
        Exit Sub
    End If
    If CInt(tCantidad.Text) < 1 Then
        MsgBox " Ingrese una cantidad positiva a modificar.", vbExclamation, "ATENCIÓN"
        tCantidad.SetFocus
        Exit Sub
    End If
    If cArticulo.ListCount = 0 Or cEstadoViejo.ListIndex = -1 Or cLocal.ListIndex = -1 Then
            MsgBox "Los datos de consulta no fueron completados.", vbExclamation, "ATENCIÓN"
            Exit Sub
    End If
    
    If opOpciones(0).Value Then 'Traslados
        Cons = "Select TraCodigo, Traslado = TraCodigo, Cantidad = RTrCantidad From Traspaso, RenglonTraspaso" _
            & " Where TraLocalIntermedio = " & CInt(Mid(cLocal.ItemData(cLocal.ListIndex), 2, 6)) _
            & " And RTrArticulo = " & cArticulo.ItemData(0) _
            & " And RTrEstado = " & cEstadoViejo.ItemData(cEstadoViejo.ListIndex) _
            & " And TraFechaEntregado = Null" _
            & " And TraCodigo = RTrTraspaso"
    Else
        Cons = "Select ReECodImpresion, Código = ReECodImpresion, ReECantidadEntregada as 'Cantidad Entregada'   From RenglonEntrega" _
            & " Where ReEArticulo = " & cArticulo.ItemData(0) _
            & " And ReECamion = " & CInt(Mid(cLocal.ItemData(cLocal.ListIndex), 2, 6)) _
            & " And ReEEstado = " & cEstadoViejo.ItemData(cEstadoViejo.ListIndex) _
            & " And ReECantidadEntregada >= " & tCantidad.Text
    End If
    Dim objLista As New clsListadeAyuda
    Dim aCodigo As Long
    Screen.MousePointer = 11
    If objLista.ActivarAyuda(cBase, Cons, 4500, 1, "Ayuda") > 0 Then
        aCodigo = objLista.RetornoDatoSeleccionado(0)
    End If
    Me.Refresh
    Screen.MousePointer = 0
    'aCodigo = ObjLista.ValorSeleccionado
    Set objLista = Nothing
    If aCodigo > 0 Then
        If opOpciones(0).Value Then
            tCodigoTraspaso.Text = aCodigo
            Foco tCodigoTraspaso
        Else
            tCodigoImpresion.Text = aCodigo
            Foco tCodigoImpresion
        End If
    End If
    Exit Sub
    
ErrBA:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error inesperado.", Err.Description
End Sub

Private Sub cArticulo_GotFocus()
    cArticulo.SelStart = 0
    cArticulo.SelLength = Len(cArticulo.Text)
    Status.SimpleText = " Ingrese el código o nombre del artículo."
End Sub
Private Sub BuscoArticuloPorNombre()
On Error GoTo ErrBAN
Dim aCodigo As Long: aCodigo = 0

    Screen.MousePointer = vbHourglass
    Cons = "Select ArtCodigo, Código = ArtCodigo, Artículo = ArtNombre from Articulo" _
        & " Where ArtNombre LIKE '" & Replace(cArticulo.Text, " ", "%") & "%'" _
        & " Order by ArtNombre"
    cArticulo.Clear
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No existe un nombre de artículo con esas características.", vbInformation, "ATENCIÓN"
    Else
        RsAux.MoveNext
        If RsAux.EOF Then
            RsAux.MoveFirst
            aCodigo = RsAux(0)
            RsAux.Close
        Else
            RsAux.Close
            Dim objLista As New clsListadeAyuda
            If objLista.ActivarAyuda(cBase, Cons, 5000, 1, "Lista Artículos") Then
                aCodigo = objLista.RetornoDatoSeleccionado(0)
            End If
            Set objLista = Nothing       'Destruyo la clase.
        End If
        If aCodigo > 0 Then BuscoArticuloCodigo aCodigo
    End If
    Screen.MousePointer = 0
    Exit Sub
    
ErrBAN:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub BuscoArticuloCodigo(aCodigo As Long)
On Error GoTo ErrBAC
    Screen.MousePointer = 11
    Cons = "Select * From Articulo Where ArtCodigo = " & aCodigo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    cArticulo.Clear
    If Not RsAux.EOF Then
        cArticulo.AddItem Trim(RsAux!ArtNombre)
        cArticulo.ItemData(cArticulo.NewIndex) = RsAux!ArtID
        cArticulo.ListIndex = 0
    End If
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub
ErrBAC:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el artículo por código.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub cArticulo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And Trim(cArticulo.Text) <> vbNullString Then
        If Not IsNumeric(cArticulo.Text) Then
            BuscoArticuloPorNombre
        Else
            BuscoArticuloCodigo cArticulo.Text
        End If
        If cArticulo.ListIndex > -1 Then Foco tCantidad
    ElseIf KeyAscii = vbKeyReturn And lvArticulo.ListItems.Count > 0 Then
        lvArticulo.SetFocus
    End If

End Sub
Private Sub cArticulo_LostFocus()
    Status.SimpleText = vbNullString
End Sub

Private Sub cEstadoNuevo_GotFocus()
    cEstadoNuevo.SelStart = 0
    cEstadoNuevo.SelLength = Len(cEstadoNuevo.Text)
    Status.SimpleText = " Seleccione el nuevo estado a pasar el artículo ."
End Sub

Private Sub cEstadoNuevo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn And cEstadoNuevo.ListIndex > -1 Then
        If cEstadoViejo.ItemData(cEstadoViejo.ListIndex) = cEstadoNuevo.ItemData(cEstadoNuevo.ListIndex) Then
            MsgBox " Selecciono el mismo estado.", vbInformation, "ATENCIÓN"
            cEstadoNuevo.SetFocus
            Exit Sub
        End If
        If opOpciones(0).Enabled Or opOpciones(1).Enabled Then
            If opOpciones(0).Enabled Then
                opOpciones(0).SetFocus
            Else
                opOpciones(1).SetFocus
            End If
        Else
            If BuscoStockLocal Then InsertoRenglon
        End If
    End If
End Sub

Private Sub cEstadoNuevo_LostFocus()
    Status.SimpleText = vbNullString
End Sub
Private Sub cEstadoViejo_GotFocus()
    cEstadoViejo.SelStart = 0
    cEstadoViejo.SelLength = Len(cEstadoViejo.Text)
    Status.SimpleText = " Seleccione el estado original del artículo ."
End Sub
Private Sub cEstadoViejo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cEstadoViejo.ListIndex > -1 Then
        If BuscoStockLocal Then
            If CInt(Mid(cLocal.ItemData(cLocal.ListIndex), 1, 1)) = 9 Then
                BuscoArticuloEnvioTraspaso
                cEstadoNuevo.SetFocus
            Else
                cEstadoNuevo.SetFocus
            End If
        Else
            cEstadoViejo.SetFocus
        End If
    End If
End Sub

Private Sub cEstadoViejo_LostFocus()
    cEstadoViejo.SelLength = 0
    Status.SimpleText = vbNullString
End Sub

Private Sub cLocal_Click()
    lvArticulo.ListItems.Clear
    LimpioDatosIngreso
End Sub

Private Sub cLocal_GotFocus()
    cLocal.SelStart = 0
    cLocal.SelLength = Len(cLocal.Text)
    Status.SimpleText = " Seleccione un Local."
End Sub
Private Sub cLocal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cLocal.ListIndex > -1 Then
            HabilitoCamion
            cArticulo.SetFocus
        Else
            MsgBox "El ingreso del local es obligatorio.", vbExclamation, "ATENCIÓN"
            cLocal.SetFocus
        End If
    End If
End Sub
Private Sub cLocal_LostFocus()
    cLocal.SelLength = 0
    Status.SimpleText = vbNullString
End Sub

Private Sub Form_Activate()
    DoEvents
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad
    CargoLocales
    CargoEstado
    DeshabilitoIngreso
    SetearLView lvValores.Grilla Or lvValores.FullRow, lvArticulo
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
End Sub

Private Sub Label1_Click()
    Foco cLocal
End Sub
Private Sub Label10_Click()
    If lvArticulo.ListItems.Count > 0 And lvArticulo.Enabled Then
        lvArticulo.ListItems(1).Selected = True
        lvArticulo.SetFocus
    End If
End Sub
Private Sub Label2_Click()
    Foco tFecha
End Sub
Private Sub Label3_Click()
    Foco cArticulo
End Sub
Private Sub Label4_Click()
    Foco tCantidad
End Sub
Private Sub Label5_Click()
    Foco cEstadoViejo
End Sub
Private Sub Label6_Click()
    Foco cEstadoNuevo
End Sub
Private Sub Label7_Click()
    Foco tComentario
End Sub
Private Sub Label8_Click()
    Foco tUsuario
End Sub

Private Sub lvArticulo_GotFocus()
    Status.SimpleText = " Lista de artículos. [Supr] - elimina"
End Sub

Private Sub lvArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            tComentario.SetFocus
        Case vbKeyDelete
            If lvArticulo.ListItems.Count > 0 Then lvArticulo.ListItems.Remove lvArticulo.SelectedItem.Index
    End Select
End Sub

Private Sub lvArticulo_LostFocus()
    Status.SimpleText = vbNullString
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
Private Sub MnuVolver_Click()
    Unload Me
End Sub
Private Sub AccionNuevo()
      
    'Habilito y Desabilito Botones
    Botones False, False, False, True, True, Toolbar1, Me
    HabilitoIngreso
    tFecha.Text = Format(Date, "d-Mmm-yyyy")
    tFecha.SetFocus
    
End Sub
Private Sub AccionGrabar()

    If ValidoDatos Then
        If MsgBox("¿Confirma almacenar los datos ingresados?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then GraboDatos
    End If
    
End Sub

Private Sub AccionCancelar()
    Screen.MousePointer = vbHourglass
    DeshabilitoIngreso
    Botones True, False, False, False, False, Toolbar1, Me
    Screen.MousePointer = vbDefault
End Sub

Private Sub opOpciones_Click(Index As Integer)

    If opOpciones(1).Value Then
        tCodigoImpresion.Enabled = True
        tCodigoImpresion.BackColor = Obligatorio
    Else
        tCodigoImpresion.Enabled = False
        tCodigoImpresion.Text = vbNullString
        tCodigoImpresion.BackColor = Inactivo
    End If
    
    If opOpciones(0).Value Then
        tCodigoTraspaso.Enabled = True
        tCodigoTraspaso.BackColor = Obligatorio
    Else
        tCodigoTraspaso.Enabled = False
        tCodigoTraspaso.Text = vbNullString
        tCodigoTraspaso.BackColor = Inactivo
    End If
    
End Sub

Private Sub opOpciones_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If opOpciones(0).Value Then
            tCodigoTraspaso.SetFocus
        Else
            tCodigoImpresion.SetFocus
        End If
    End If
End Sub

Private Sub tCantidad_GotFocus()
    tCantidad.SelStart = 0
    tCantidad.SelLength = Len(tCantidad.Text)
    Status.SimpleText = " Ingrese la cantidad de artículos."
End Sub

Private Sub tCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(tCantidad.Text) <> vbNullString Then
        If IsNumeric(tCantidad.Text) Then
            If cEstadoViejo.ListIndex = -1 Then BuscoCodigoEnCombo cEstadoViejo, CInt(paEstadoArticuloEntrega)
            cEstadoViejo.SetFocus
        Else
            MsgBox " El formato ingresado no es numérico.", vbExclamation, "ATENCIÓN"
            tCantidad.SetFocus
        End If
    End If
End Sub

Private Sub tCodigoImpresion_GotFocus()
    tCodigoImpresion.SelStart = 0
    tCodigoImpresion.SelLength = Len(tCodigoImpresion.Text)
    Status.SimpleText = " Ingrese el código de impresión."
End Sub

Private Sub tCodigoImpresion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tCodigoImpresion.Text) Then
            InsertoRenglon
        Else
            MsgBox "No se ingreso un número.", vbExclamation, "ATENCIÓN"
            tCodigoImpresion.SelStart = 0
            tCodigoImpresion.SelLength = Len(tCodigoImpresion.Text)
            tCodigoImpresion.SetFocus
        End If
    End If
End Sub

Private Sub tCodigoImpresion_LostFocus()
    Status.SimpleText = vbNullString
End Sub


Private Sub tCodigoTraspaso_GotFocus()
    tCodigoTraspaso.SelStart = 0
    tCodigoTraspaso.SelLength = Len(tCodigoTraspaso.Text)
    Status.SimpleText = " Ingrese el código de traspaso."
End Sub

Private Sub tCodigoTraspaso_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tCodigoTraspaso.Text) Then
            InsertoRenglon
        Else
            MsgBox "No se ingreso un número.", vbExclamation, "ATENCIÓN"
            tCodigoTraspaso.SelStart = 0
            tCodigoTraspaso.SelLength = Len(tCodigoTraspaso.Text)
            tCodigoTraspaso.SetFocus
        End If
    End If
End Sub

Private Sub tCodigoTraspaso_LostFocus()
    Status.SimpleText = vbNullString
End Sub
Private Sub tComentario_GotFocus()
    tComentario.SelStart = 0
    tComentario.SelLength = Len(tComentario.Text)
    Status.SimpleText = " Ingrese un comentario."
End Sub
Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tUsuario.SetFocus
End Sub
Private Sub tFecha_GotFocus()
    tFecha.SelStart = 0
    tFecha.SelLength = Len(tFecha.Text)
    Status.SimpleText = " Ingrese una fecha."
End Sub
Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsDate(tFecha.Text) Then
            cLocal.SetFocus
        Else
            MsgBox " La fecha ingresada no es correcta.", vbExclamation, "ATENCIÓN"
            tFecha.SetFocus
        End If
    End If
End Sub

Private Sub tFecha_LostFocus()
    If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, "d-Mmm-yyyy")
    Status.SimpleText = vbNullString
End Sub

Private Sub tFecha_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.SimpleText = " Ingrese una fecha."
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    
    Select Case Button.Key
        
        Case "nuevo": AccionNuevo
        
        Case "grabar"
            AccionGrabar
        
        Case "cancelar"
            AccionCancelar
        
        Case "salir"
            Unload Me
            
    End Select

End Sub
Private Sub DeshabilitoIngreso()
    tFecha.Text = vbNullString
    tFecha.Enabled = False
    tFecha.BackColor = Inactivo
    cLocal.ListIndex = -1
    cLocal.Enabled = False
    cLocal.BackColor = Inactivo
    cArticulo.Enabled = False
    cArticulo.BackColor = Inactivo
    cArticulo.Clear
    tCantidad.Text = vbNullString
    tCantidad.Enabled = False
    tCantidad.BackColor = Inactivo
    cEstadoViejo.Enabled = False
    cEstadoViejo.BackColor = Inactivo
    cEstadoViejo.ListIndex = -1
    cEstadoNuevo.Enabled = False
    cEstadoNuevo.BackColor = Inactivo
    cEstadoNuevo.ListIndex = -1
    tComentario.BackColor = Inactivo
    tComentario.Enabled = False
    tComentario.Text = vbNullString
    DeshabilitoStockCamion
    lvArticulo.Enabled = False
    lvArticulo.ListItems.Clear
    tUsuario.Enabled = False
    tUsuario.BackColor = Inactivo
    tUsuario.Text = vbNullString
    tUsuario.Tag = vbNullString
End Sub
Private Sub HabilitoIngreso()
    tFecha.Enabled = True
    tFecha.BackColor = Obligatorio
    cLocal.Enabled = True
    cLocal.BackColor = Obligatorio
    cArticulo.Enabled = True
    cArticulo.BackColor = Obligatorio
    cArticulo.Clear
    tCantidad.Text = vbNullString
    tCantidad.Enabled = True
    tCantidad.BackColor = Obligatorio
    cEstadoViejo.Enabled = True
    cEstadoViejo.BackColor = Obligatorio
    cEstadoNuevo.Enabled = True
    cEstadoNuevo.BackColor = Obligatorio
    tComentario.BackColor = Blanco
    tComentario.Enabled = True
    tComentario.Text = vbNullString
    lvArticulo.Enabled = True
    lvArticulo.ListItems.Clear
    tUsuario.Enabled = True
    tUsuario.BackColor = Obligatorio
    tUsuario.Text = vbNullString
    tUsuario.Tag = vbNullString
End Sub
Private Sub DeshabilitoStockCamion()
    bAyuda.Enabled = False
    opOpciones(0).Enabled = False
    opOpciones(0).Value = False
    opOpciones(1).Enabled = False
    opOpciones(1).Value = False
    tCodigoImpresion.Enabled = False
    tCodigoImpresion.BackColor = Inactivo
    tCodigoImpresion.Text = vbNullString
    tCodigoTraspaso.Enabled = False
    tCodigoTraspaso.BackColor = Inactivo
    tCodigoTraspaso.Text = vbNullString
End Sub
Private Sub HabilitoStockCamion()
    bAyuda.Enabled = True
    opOpciones(0).Enabled = True
    opOpciones(0).Value = True
    opOpciones(1).Enabled = True
    opOpciones(0).Value = False
End Sub
Private Sub CargoLocales()
On Error GoTo ErrCL
    cLocal.Clear
    If prmEsp Then
        Cons = "Select * From Local Order by LocNombre"
    Else
        Cons = "Select * from local where LocCodigo = " & paCodigoDeSucursal
    End If
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        cLocal.AddItem Trim(RsAux("LocNombre"))
        If RsAux!LocTipo = 1 Then
            cLocal.ItemData(cLocal.NewIndex) = CLng("9" & RsAux("LocCodigo"))
        Else
            cLocal.ItemData(cLocal.NewIndex) = CLng("1" & RsAux("LocCodigo"))
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    Exit Sub
ErrCL:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los locales."
    Screen.MousePointer = vbDefault
End Sub
Private Sub CargoEstado()
On Error GoTo ErrCE
    Cons = "Select EsMCodigo, EsMAbreviacion From EstadoMercaderia " _
        & " Order by EsMAbreviacion"
    CargoCombo Cons, cEstadoNuevo, ""
    CargoCombo Cons, cEstadoViejo, ""
    Exit Sub
ErrCE:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al cargar los Estados."
End Sub
Private Sub HabilitoCamion()
    If CInt(Mid(cLocal.ItemData(cLocal.ListIndex), 1, 1)) = 9 Then
        HabilitoStockCamion
    Else
        DeshabilitoStockCamion
    End If
End Sub
Private Sub BuscoArticuloEnvioTraspaso()
On Error GoTo ErrBAET

    Screen.MousePointer = vbHourglass
    Cons = "Select * From Traspaso, RenglonTraspaso" _
        & " Where TraLocalIntermedio = " & CInt(Mid(cLocal.ItemData(cLocal.ListIndex), 2, 6)) _
        & " And RTrArticulo = " & cArticulo.ItemData(0) _
        & " And RTrCantidad >= " & tCantidad.Text _
        & " And RTrEstado = " & cEstadoViejo.ItemData(cEstadoViejo.ListIndex) _
        & " And TraFechaEntregado = Null" _
        & " And TraCodigo = RTrTraspaso"
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If RsAux.EOF Then
        opOpciones(0).Enabled = False
    Else
        opOpciones(0).Enabled = True
    End If
    RsAux.Close
    
    If Not HayEnEnvio(0) Then
        opOpciones(1).Enabled = False
    Else
        opOpciones(1).Enabled = True
    End If
    
    If opOpciones(0).Enabled Then
        opOpciones(0).Value = True
    ElseIf opOpciones(1).Enabled Then
        opOpciones(1).Value = True
    End If
         
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrBAET:
    clsGeneral.OcurrioError "Ocurrió un error al buscar envíos y traspasos."
End Sub
Private Function BuscoStockLocal() As Boolean
On Error GoTo ErrBSL

    BuscoStockLocal = False
    If cLocal.ListIndex = -1 Or cEstadoViejo.ListIndex = -1 Then Exit Function
        
    Screen.MousePointer = vbHourglass
    Cons = "Select StLCantidad from StockLocal" _
        & " Where StLLocal = " & CInt(Mid(cLocal.ItemData(cLocal.ListIndex), 2, 6)) _
        & " And StLArticulo = " & cArticulo.ItemData(0) _
        & " And StLEstado = " & cEstadoViejo.ItemData(cEstadoViejo.ListIndex)
        
    If CInt(Mid(cLocal.ItemData(cLocal.ListIndex), 1, 1)) = 9 Then
        Cons = Cons & " And StLTipoLocal = " & TipoLocal.Camion
    Else
        Cons = Cons & " And StLTipoLocal = " & TipoLocal.Deposito
    End If
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If RsAux.EOF Then
        BuscoStockLocal = False
        MsgBox "El local no posee el artículo con ese estado.", vbInformation, "ATENCIÓN"
    Else
        If RsAux!StLCantidad < CInt(tCantidad.Text) Then
            BuscoStockLocal = False
            MsgBox "El local no posee tantos artículos con ese estado, no puede ingresar este renglón.", vbInformation, "ATENCIÓN"
        Else
            BuscoStockLocal = True
        End If
    End If
    RsAux.Close
    
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrBSL:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el stock del local."
End Function

Private Sub InsertoRenglon()
Dim Msg As String
On Error GoTo ErrControl

    Screen.MousePointer = vbHourglass
    If cArticulo.ListIndex = -1 Then
        MsgBox "No hay seleccionado un artículo.", vbExclamation, "ATENCIÓN"
        cArticulo.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If Not IsNumeric(tCantidad.Text) Then
        MsgBox "La cantidad ingresada no es correcta.", vbInformation, "ATENCIÓN"
        tCantidad.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If CInt(tCantidad.Text) < 1 Then
        MsgBox "La cantidad ingresada no es correcta.", vbInformation, "ATENCIÓN"
        tCantidad.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If cEstadoViejo.ListIndex = -1 Then
        MsgBox "No hay seleccionado un estado.", vbInformation, "ATENCIÓN"
        cEstadoViejo.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If cEstadoNuevo.ListIndex = -1 Then
        MsgBox "Debe seleccionar el nuevo estado.", vbInformation, "ATENCIÓN"
        cEstadoNuevo.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If cEstadoViejo.ItemData(cEstadoViejo.ListIndex) = cEstadoNuevo.ItemData(cEstadoNuevo.ListIndex) Then
        MsgBox " Selecciono el mismo estado.", vbInformation, "ATENCIÓN"
        cEstadoNuevo.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If opOpciones(1).Value Then
        If Trim(tCodigoImpresion.Text) = vbNullString Then
            MsgBox "No hay seleccionado un código de impresión.", vbInformation, "ATENCIÓN"
            tCodigoImpresion.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If Not IsNumeric(tCodigoImpresion.Text) Then
            MsgBox "El código de impresión seleccionado no es correcto.", vbInformation, "ATENCIÓN"
            tCodigoImpresion.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If Not HayEnEnvio(CInt(tCodigoImpresion.Text)) Then
            MsgBox "El código de impresión seleccionado no cumple con las condiciones, verifique.", vbInformation, "ATENCIÓN"
            tCodigoImpresion.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    ElseIf opOpciones(0).Value Then
        If Trim(tCodigoTraspaso.Text) = vbNullString Then
            MsgBox "No hay seleccionado un código de traspaso.", vbInformation, "ATENCIÓN"
            tCodigoTraspaso.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If Not IsNumeric(tCodigoTraspaso.Text) Then
            MsgBox "El código de traslado seleccionado no es correcto.", vbInformation, "ATENCIÓN"
            tCodigoTraspaso.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        Cons = "Select * From Traspaso, RenglonTraspaso" _
            & " Where TraLocalIntermedio = " & CInt(Mid(cLocal.ItemData(cLocal.ListIndex), 2, 6)) _
            & " And RTrArticulo = " & cArticulo.ItemData(0) _
            & " And RTrEstado = " & cEstadoViejo.ItemData(cEstadoViejo.ListIndex) _
            & " And RTrCantidad >= " & tCantidad.Text _
            & " And TraCodigo = " & tCodigoTraspaso.Text _
            & " And TraFechaEntregado = Null" _
            & " And TraCodigo = RTrTraspaso"
            
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If RsAux.EOF Then
            RsAux.Close
            MsgBox "El traspaso seleccionado no posee esas condiciones.", vbExclamation, "ATENCIÓN"
            tCodigoTraspaso.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            RsAux.Close
        End If
    End If
    
    'Si llegue aca es porque puedo insertar.
    On Error GoTo ErrInserto
    Msg = " Ya se ingreso ese artículo con las mismas condiciones."
    'Clave Articulo + estado anterior + estado nuevo
    ' A la clave lleva adelante si es envio o traspaso, para las sucursales se considera como envío.
    If opOpciones(0).Value Then
        Set itmX = lvArticulo.ListItems.Add(, "T" & cArticulo.ItemData(0) & "V" & cEstadoViejo.ItemData(cEstadoViejo.ListIndex) & "N" & cEstadoNuevo.ItemData(cEstadoNuevo.ListIndex))
    Else
        Set itmX = lvArticulo.ListItems.Add(, "E" & cArticulo.ItemData(0) & "V" & cEstadoViejo.ItemData(cEstadoViejo.ListIndex) & "N" & cEstadoNuevo.ItemData(cEstadoNuevo.ListIndex))
    End If
    itmX.Text = tCantidad.Text
    itmX.SubItems(1) = Trim(cArticulo.Text)
    itmX.SubItems(2) = Trim(cEstadoViejo.Text)
    itmX.SubItems(3) = Trim(cEstadoNuevo.Text)
    If Mid(cLocal.ItemData(cLocal.ListIndex), 1, 1) = 9 Then
        If opOpciones(0).Value Then
            itmX.Tag = tCodigoTraspaso.Text
            itmX.SubItems(4) = "T" & Trim(tCodigoTraspaso.Text)
        Else
            itmX.Tag = tCodigoImpresion.Text
            itmX.SubItems(4) = "E" & Trim(tCodigoImpresion.Text)
        End If
    End If
    LimpioDatosIngreso
    cArticulo.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrControl:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al controlar los datos."
    Exit Sub
    
ErrInserto:
    Screen.MousePointer = vbDefault
    If Msg = vbNullString Then
        clsGeneral.OcurrioError Err.Description
    Else
        clsGeneral.OcurrioError Msg
    End If
End Sub
Private Function HayEnEnvio(iCodImpresion As Integer) As Boolean

    Cons = "Select * From RenglonEntrega" _
        & " Where ReEArticulo = " & cArticulo.ItemData(0) _
        & " And ReECamion = " & CInt(Mid(cLocal.ItemData(cLocal.ListIndex), 2, 6)) _
        & " And ReECantidadEntregada >= " & tCantidad.Text _
        & " And ReEEstado = " & cEstadoViejo.ItemData(cEstadoViejo.ListIndex)
        
    If iCodImpresion > 0 Then Cons = Cons & " And ReECodImpresion = " & iCodImpresion

    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If RsAux.EOF Then
        HayEnEnvio = False
    Else
        HayEnEnvio = True
    End If
    RsAux.Close
    
End Function
Private Sub LimpioDatosIngreso()

    cArticulo.Clear
    tCantidad.Text = vbNullString
    cEstadoViejo.Text = ""
    cEstadoNuevo.Text = ""
    tCodigoImpresion.Text = vbNullString
    tCodigoTraspaso.Text = vbNullString
    
End Sub
Private Sub tUsuario_GotFocus()
    tUsuario.SelStart = 0
    tUsuario.SelLength = Len(tUsuario.Text)
    Status.SimpleText = " Ingrese su código de Usuario."
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)
    tUsuario.Tag = vbNullString
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tUsuario.Text) Then
            tUsuario.Tag = BuscoUsuarioDigito(CInt(tUsuario.Text), True)
            If CInt(tUsuario.Tag) > 0 Then
                AccionGrabar
            Else
                tUsuario.Tag = vbNullString
            End If
        Else
            MsgBox "El formato del código de usuario no es numérico.", vbExclamation, "ATENCIÓN"
            tUsuario.SetFocus
        End If
    End If
End Sub

Private Function ValidoDatos() As Boolean

    ValidoDatos = True
    
    If lvArticulo.ListItems.Count = 0 Then
        MsgBox "No hay artículos ingresados.", vbExclamation, "ATENCIÓN"
        ValidoDatos = False
        Exit Function
    End If
    
    If tUsuario.Tag = vbNullString Then
        MsgBox "Debe ingresar su código de usuario.", vbExclamation, "ATENCIÓN"
        tUsuario.SetFocus
        ValidoDatos = False
        Exit Function
    End If
    
    If Not clsGeneral.TextoValido(tComentario.Text) Then
        MsgBox "Se ingreso un carácter no válido en el comentario.", vbExclamation, "ATENCIÓN"
        tComentario.SetFocus
        ValidoDatos = False
        Exit Function
    End If
    
End Function
Private Sub GraboDatos()
On Error GoTo ErrGD
Dim aCod As Long
    
    FechaDelServidor
    cBase.BeginTrans
    On Error GoTo errResumo
    
    Cons = "INSERT INTO ControlMercaderia (CMeTipoLocal, CMeLocal, CMeFecha, CMeTipo, CMeComentario, CMeUsuario)" _
        & " Values ("
        
    If CInt(Mid(cLocal.ItemData(cLocal.ListIndex), 1, 1)) = 9 Then
        Cons = Cons & TipoLocal.Camion
    Else
        Cons = Cons & TipoLocal.Deposito
    End If
    
    Cons = Cons & ", " & CInt(Mid(cLocal.ItemData(cLocal.ListIndex), 2, 6)) _
        & ", '" & Format(gFechaServidor, FormatoFH) & "', " & TipoControlMercaderia.CambioEstado
        
    If Trim(tComentario.Text) = vbNullString Then
        Cons = Cons & ", Null, " & tUsuario.Tag & ")"
    Else
        Cons = Cons & ", '" & tComentario.Text & "', " & tUsuario.Tag & ")"
    End If
    cBase.Execute (Cons)
    
    Cons = "Select Max(CMeCodigo) From ControlMercaderia Where CMeTipo = " & TipoControlMercaderia.CambioEstado
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    aCod = RsAux(0)
    RsAux.Close
    For Each itmX In lvArticulo.ListItems
        If CInt(Mid(cLocal.ItemData(cLocal.ListIndex), 1, 1)) = 9 Then
            InsertoRenglonCamion aCod
        Else
            InsertoRenglonLocal aCod
        End If
    Next
    
    cBase.CommitTrans
    
    AccionCancelar
    Exit Sub

ErrGD:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al iniciar la transacción."
    Exit Sub
    
errResumo:
    Resume Relajo
    
Relajo:
    Screen.MousePointer = vbDefault
    cBase.RollbackTrans
    If Msg = vbNullString Then
        clsGeneral.OcurrioError Err.Description
    Else
        clsGeneral.OcurrioError Msg
    End If
    Exit Sub
    
End Sub
Private Sub InsertoRenglonLocal(IDContMerc As Long)
Dim sAfectaNuevo As Boolean
Dim sAfectaViejo As Boolean

    Cons = "Select * From StockLocal " _
        & " Where StLTipoLocal = " & TipoLocal.Deposito _
        & " And StlLocal = " & Mid(cLocal.ItemData(cLocal.ListIndex), 2, 6) _
        & " And StLArticulo = " & Mid(itmX.Key, 2, InStr(itmX.Key, "V") - 2) _
        & " And StLEstado = " & Mid(itmX.Key, InStr(itmX.Key, "V") + 1, InStr(itmX.Key, "N") - InStr(itmX.Key, "V") - 1)
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        Msg = "El local no posee más artículos " & Trim(itmX.SubItems(1))
        RsAux.Close
        RsAux.Edit
    Else
        If RsAux!StLCantidad < CInt(itmX.Text) Then
            Msg = "El local no posee tantos artículos " & Trim(itmX.SubItems(1))
            RsAux.Close
            RsAux.Edit
        ElseIf RsAux!StLCantidad = CInt(itmX.Text) Then
            RsAux.Delete
            RsAux.Close
        Else
            RsAux.Edit
            RsAux!StLCantidad = RsAux!StLCantidad - CInt(itmX.Text)
            RsAux.Update
            RsAux.Close
        End If
    End If
    MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, Mid(cLocal.ItemData(cLocal.ListIndex), 2, 6), Mid(itmX.Key, 2, InStr(itmX.Key, "V") - 2), CInt(itmX.Text), Mid(itmX.Key, InStr(itmX.Key, "V") + 1, InStr(itmX.Key, "N") - InStr(itmX.Key, "V") - 1), -1, TipoDocumento.CambioEstadoMercaderia, IDContMerc
    
    Cons = "Select * From StockLocal " _
        & " Where StLTipoLocal = " & TipoLocal.Deposito _
        & " And StlLocal = " & Mid(cLocal.ItemData(cLocal.ListIndex), 2, 6) _
        & " And StLArticulo = " & Mid(itmX.Key, 2, InStr(itmX.Key, "V") - 2) _
        & " And StLEstado = " & Mid(itmX.Key, InStr(itmX.Key, "N") + 1, Len(itmX.Key))
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        RsAux.AddNew
        RsAux!StLTipoLocal = TipoLocal.Deposito
        RsAux!StlLocal = Mid(cLocal.ItemData(cLocal.ListIndex), 2, 6)
        RsAux!StLArticulo = Mid(itmX.Key, 2, InStr(itmX.Key, "V") - 2)
        RsAux!StLEstado = Mid(itmX.Key, InStr(itmX.Key, "N") + 1, Len(itmX.Key))
        RsAux!StLCantidad = CInt(itmX.Text)
        RsAux.Update
        RsAux.Close
    Else
        RsAux.Edit
        RsAux!StLCantidad = RsAux!StLCantidad + CInt(itmX.Text)
        RsAux.Update
        RsAux.Close
    End If
    MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, Mid(cLocal.ItemData(cLocal.ListIndex), 2, 6), Mid(itmX.Key, 2, InStr(itmX.Key, "V") - 2), CInt(itmX.Text), Mid(itmX.Key, InStr(itmX.Key, "N") + 1, Len(itmX.Key)), 1, TipoDocumento.CambioEstadoMercaderia, IDContMerc
    
    'El estado anterior era sano, a rec., etc., entonces le quito la cantidad a ese estado.
    MarcoMovimientoStockTotal Mid(itmX.Key, 2, InStr(itmX.Key, "V") - 2), TipoEstadoMercaderia.Fisico, Mid(itmX.Key, InStr(itmX.Key, "V") + 1, InStr(itmX.Key, "N") - InStr(itmX.Key, "V") - 1), CInt(itmX.Text), -1
    MarcoMovimientoStockTotal Mid(itmX.Key, 2, InStr(itmX.Key, "V") - 2), TipoEstadoMercaderia.Fisico, Mid(itmX.Key, InStr(itmX.Key, "N") + 1, Len(itmX.Key)), CInt(itmX.Text), 1
    
    
End Sub
Private Sub InsertoRenglonCamion(IDContMerc As Long)
Dim sAfectaNuevo As Boolean
Dim sAfectaViejo As Boolean

    'Veo si afecta el stock físico.
    'Cons = "Select EsMBajaStockTotal From EstadoMercaderia" _
        & " Where EsMCodigo = " & Mid(itmX.Key, InStr(itmX.Key, "N") + 1, Len(itmX.Key))
    
    'Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    'If RsAux.EOF Then
    '    Msg = "No se encontró los datos del nuevo estado, reintente."
    '    RsAux.Close
    '    RsAux.Edit
    'Else
    '    sAfectaNuevo = False
    '    If RsAux!EsMBajaStockTotal = 1 Then
    '       sAfectaNuevo = True
    '    End If
    '    RsAux.Close
    'End If
    
    'Cons = "Select EsMBajaStockTotal From EstadoMercaderia" _
        & " Where EsMCodigo = " & Mid(itmX.Key, InStr(itmX.Key, "V") + 1, InStr(itmX.Key, "N") - InStr(itmX.Key, "V") - 1)
    
    'Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    'If RsAux.EOF Then
    '    Msg = "No se encontró los datos del nuevo estado, reintente."
    '    RsAux.Close
    '    RsAux.Edit
    'Else
    '    sAfectaViejo = False
    '    If RsAux!EsMBajaStockTotal = 1 Then
    '       sAfectaViejo = True
    '    End If
    '    RsAux.Close
'    End If

    
    If Mid(itmX.SubItems(4), 1, 1) = "T" Then 'Es traspaso, verifico si el mismo tiene cantidad.
        
        Cons = "Select * From Traspaso, RenglonTraspaso" _
            & " Where TraLocalIntermedio = " & CInt(Mid(cLocal.ItemData(cLocal.ListIndex), 2, 6)) _
            & " And RTrArticulo = " & Mid(itmX.Key, 2, InStr(itmX.Key, "V") - 2) _
            & " And RTrEstado = " & Mid(itmX.Key, InStr(itmX.Key, "V") + 1, InStr(itmX.Key, "N") - InStr(itmX.Key, "V") - 1) _
            & " And TraCodigo = " & itmX.Tag _
            & " And TraCodigo = RTrTraspaso"
            
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If RsAux.EOF Then
            Msg = "El traspaso " & itmX.Tag & " fue eliminado o modificado, verifique."
            RsAux.Close
            RsAux.Edit
        Else
            If RsAux!TraFechaEntregado <> Null Then
                Msg = "El traspaso " & itmX.Tag & " fue actualizado como entregado, verifique."
                RsAux.Close
                RsAux.Edit
            Else
                If RsAux!RTrCantidad < CInt(itmX.Text) Then
                    Msg = "En el traspaso no quedan tantos artículos " & itmX.SubItems(1)
                    RsAux.Close
                    RsAux.Edit
                Else
                    Cons = "Update RenglonTraspaso Set RTrCantidad = RTrCantidad -" & CInt(itmX.Text) _
                        & " Where RTrTraspaso = " & itmX.Tag _
                        & " And RTrArticulo = " & Mid(itmX.Key, 2, InStr(itmX.Key, "V") - 2) _
                        & " And RTrEstado = " & Mid(itmX.Key, InStr(itmX.Key, "V") + 1, InStr(itmX.Key, "N") - InStr(itmX.Key, "V") - 1)
                    
                    cBase.Execute (Cons)
                    
                    If cBase.RowsAffected = 0 Then
                        Msg = "No se actualizó la tabla RenglonTraspaso."
                        RsAux.Close
                        RsAux.Edit
                    End If
                    RsAux.Close
                End If
                
               ' If Not sAfectaNuevo Then
                    'Ahora inserto o updateo para el nuevo estado.
                    Cons = "Select * From RenglonTraspaso" _
                        & " Where RTrTraspaso = " & itmX.Tag _
                        & " And RTrArticulo = " & Mid(itmX.Key, 2, InStr(itmX.Key, "V") - 2) _
                        & " And RTrEstado = " & Mid(itmX.Key, InStr(itmX.Key, "N") + 1, Len(itmX.Key))
                        
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    
                    If RsAux.EOF Then
                        RsAux.AddNew
                        RsAux!RTrTraspaso = itmX.Tag
                        RsAux!RTrArticulo = Mid(itmX.Key, 2, InStr(itmX.Key, "V") - 2)
                        RsAux!RTrEstado = Mid(itmX.Key, InStr(itmX.Key, "N") + 1, Len(itmX.Key))
                        RsAux!RTrCantidad = CInt(itmX.Text)
                        RsAux.Update
                    Else
                        RsAux.Edit
                        RsAux!RTrCantidad = RsAux!RTrCantiad + CInt(itmX.Text)
                        RsAux.Update
                    End If
                    RsAux.Close
               ' End If
            End If
        End If
    Else
        'Es por envio
        Cons = "Select * From RenglonEntrega" _
            & " Where ReEArticulo = " & Mid(itmX.Key, 2, InStr(itmX.Key, "V") - 2) _
            & " And ReECamion = " & CInt(Mid(cLocal.ItemData(cLocal.ListIndex), 2, 6)) _
            & " And ReEEstado = " & Mid(itmX.Key, InStr(itmX.Key, "V") + 1, InStr(itmX.Key, "N") - InStr(itmX.Key, "V") - 1) _
            & " And ReECodImpresion = " & itmX.Tag
    
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If RsAux.EOF Then
            Msg = "Otra terminal pudo eliminar o modificar el código de impresión : " & itmX.Tag & " ."
            RsAux.Close
            RsAux.Edit
        Else
            If RsAux!ReECantidadEntregada < CInt(itmX.Text) Then
                Msg = "No hay tantos artículos " & Trim(itmX.SubItems(1)) & " en el código de impresión " & itmX.Tag & " ."
                RsAux.Close
                RsAux.Edit
            ElseIf RsAux!ReECantidadEntregada > CInt(itmX.Text) Then
                RsAux.Edit
                RsAux!ReECantidadTotal = RsAux!ReECantidadTotal - CInt(itmX.Text)
                RsAux!ReECantidadEntregada = RsAux!ReECantidadEntregada - CInt(itmX.Text)
                RsAux.Update
            Else
                If paEstadoArticuloEntrega = RsAux!ReEEstado Then
                    'Nunca puedo borrar el estado sano.
                    RsAux.Edit
                    RsAux!ReECantidadTotal = RsAux!ReECantidadTotal - CInt(itmX.Text)
                    RsAux!ReECantidadEntregada = RsAux!ReECantidadEntregada - CInt(itmX.Text)
                    RsAux.Update
                Else
                    RsAux.Delete
                End If
            End If
            RsAux.Close
            
            'If Not sAfectaNuevo Then
                Cons = "Select * From RenglonEntrega" _
                    & " Where ReEArticulo = " & Mid(itmX.Key, 2, InStr(itmX.Key, "V") - 2) _
                    & " And ReECamion = " & CInt(Mid(cLocal.ItemData(cLocal.ListIndex), 2, 6)) _
                    & " And ReEEstado = " & Mid(itmX.Key, InStr(itmX.Key, "N") + 1, Len(itmX.Key)) _
                    & " And ReECodImpresion = " & itmX.Tag
                
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                
                If RsAux.EOF Then
                    RsAux.AddNew
                    RsAux!ReECodImpresion = itmX.Tag
                    RsAux!ReEArticulo = Mid(itmX.Key, 2, InStr(itmX.Key, "V") - 2)
                    RsAux!ReECamion = CInt(Mid(cLocal.ItemData(cLocal.ListIndex), 2, 6))
                    RsAux!ReEEstado = Mid(itmX.Key, InStr(itmX.Key, "N") + 1, Len(itmX.Key))
                    RsAux!ReECantidadTotal = CInt(itmX.Text)
                    RsAux!ReECantidadEntregada = CInt(itmX.Text)
                    RsAux!ReEFModificacion = Format(Now, FormatoFH)
                    RsAux.Update
                Else
                    RsAux.Edit
                    RsAux!ReECantidadTotal = RsAux!ReECantidadTotal + CInt(itmX.Text)
                    RsAux!ReECantidadEntregada = RsAux!ReECantidadEntregada + CInt(itmX.Text)
                    RsAux.Update
                End If
                RsAux.Close
            'End If
            
            Cons = "Update RenglonEntrega Set ReEFModificacion = '" & Format(Now, FormatoFH) & "'" _
                & " Where ReECodImpresion = " & itmX.Tag
            
            cBase.Execute (Cons)
        End If
    End If
    
    Cons = "Select * From StockLocal " _
        & " Where StLTipoLocal = " & TipoLocal.Camion _
        & " And StlLocal = " & Mid(cLocal.ItemData(cLocal.ListIndex), 2, 6) _
        & " And StLArticulo = " & Mid(itmX.Key, 2, InStr(itmX.Key, "V") - 2) _
        & " And StLEstado = " & Mid(itmX.Key, InStr(itmX.Key, "V") + 1, InStr(itmX.Key, "N") - InStr(itmX.Key, "V") - 1)
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        Msg = "El local no posee más artículos " & Trim(itmX.SubItems(1))
        RsAux.Close
        RsAux.Edit
    Else
        If RsAux!StLCantidad < CInt(itmX.Text) Then
            Msg = "El local no posee tantos artículos " & Trim(itmX.SubItems(1))
            RsAux.Close
            RsAux.Edit
        ElseIf RsAux!StLCantidad = CInt(itmX.Text) Then
            RsAux.Delete
            RsAux.Close
        Else
            RsAux.Edit
            RsAux!StLCantidad = RsAux!StLCantidad - CInt(itmX.Text)
            RsAux.Update
            RsAux.Close
        End If
    End If
    MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Camion, Mid(cLocal.ItemData(cLocal.ListIndex), 2, 6), Mid(itmX.Key, 2, InStr(itmX.Key, "V") - 2), CInt(itmX.Text), Mid(itmX.Key, InStr(itmX.Key, "V") + 1, InStr(itmX.Key, "N") - InStr(itmX.Key, "V") - 1), -1, TipoDocumento.CambioEstadoMercaderia, IDContMerc
    
    Cons = "Select * From StockLocal " _
        & " Where StLTipoLocal = " & TipoLocal.Camion _
        & " And StlLocal = " & Mid(cLocal.ItemData(cLocal.ListIndex), 2, 6) _
        & " And StLArticulo = " & Mid(itmX.Key, 2, InStr(itmX.Key, "V") - 2) _
        & " And StLEstado = " & Mid(itmX.Key, InStr(itmX.Key, "N") + 1, Len(itmX.Key))
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        RsAux.AddNew
        RsAux!StLTipoLocal = TipoLocal.Camion
        RsAux!StlLocal = Mid(cLocal.ItemData(cLocal.ListIndex), 2, 6)
        RsAux!StLArticulo = Mid(itmX.Key, 2, InStr(itmX.Key, "V") - 2)
        RsAux!StLEstado = Mid(itmX.Key, InStr(itmX.Key, "N") + 1, Len(itmX.Key))
        RsAux!StLCantidad = CInt(itmX.Text)
        RsAux.Update
        RsAux.Close
    Else
        RsAux.Edit
        RsAux!StLCantidad = RsAux!StLCantidad + CInt(itmX.Text)
        RsAux.Update
        RsAux.Close
    End If
    MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Camion, Mid(cLocal.ItemData(cLocal.ListIndex), 2, 6), Mid(itmX.Key, 2, InStr(itmX.Key, "V") - 2), CInt(itmX.Text), Mid(itmX.Key, InStr(itmX.Key, "N") + 1, Len(itmX.Key)), 1, TipoDocumento.CambioEstadoMercaderia, IDContMerc
 
   ' If Not sAfectaViejo Then
        'El estado anterior era sano, a rec., etc., entonces le quito la cantidad a ese estado.
        Cons = "Select * From StockTotal " _
            & " Where StTArticulo = " & Mid(itmX.Key, 2, InStr(itmX.Key, "V") - 2) _
            & " And StTTipoEstado = " & TipoEstadoMercaderia.Fisico _
            & " And StTEstado = " & Mid(itmX.Key, InStr(itmX.Key, "V") + 1, InStr(itmX.Key, "N") - InStr(itmX.Key, "V") - 1)
            
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If RsAux.EOF Then
            Msg = "El stock total para el artículo " & Trim(itmX.SubItems(1)) & " no existe."
            RsAux.Close
            RsAux.Edit
        Else
            If RsAux!StTCantidad >= CInt(itmX.Text) Then
                RsAux.Edit
                RsAux!StTCantidad = RsAux!StTCantidad - CInt(itmX.Text)
                RsAux.Update
                RsAux.Close
            Else
                Msg = "No hay tantos artículos " & Trim(itmX.SubItems(1)) & " en el stock total para ese estado."
                RsAux.Close
                RsAux.Edit
            End If
        End If
   ' End If
    
   ' If Not sAfectaNuevo Then
      Cons = "Select * From StockTotal " _
            & " Where StTArticulo = " & Mid(itmX.Key, 2, InStr(itmX.Key, "V") - 2) _
            & " And StTTipoEstado = " & TipoEstadoMercaderia.Fisico _
            & " And StTEstado = " & Mid(itmX.Key, InStr(itmX.Key, "N") + 1, Len(itmX.Key))
            
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If RsAux.EOF Then
            RsAux.AddNew
            RsAux!StTArticulo = Mid(itmX.Key, 2, InStr(itmX.Key, "V") - 2)
            RsAux!StTTipoEstado = TipoEstadoMercaderia.Fisico
            RsAux!StTEstado = Mid(itmX.Key, InStr(itmX.Key, "N") + 1, Len(itmX.Key))
            RsAux!StTCantidad = CInt(itmX.Text)
            RsAux.Update
        Else
            RsAux.Edit
            RsAux!StTCantidad = RsAux!StTCantidad + CInt(itmX.Text)
            RsAux.Update
        End If
        RsAux.Close
   ' End If
    
End Sub
