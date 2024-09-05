VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.0#0"; "AACOMBO.OCX"
Begin VB.Form frmCatDto 
   Caption         =   "Asignar Descuentos"
   ClientHeight    =   5700
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCategDto.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   11
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "refrescar"
            Object.ToolTipText     =   "Refrescar datos"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin ComctlLib.ListView lvIngresados 
      Height          =   3615
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Cat. Artículos"
         Object.Width           =   3422
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Cat. Clientes"
         Object.Width           =   2981
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Cat. Plazos"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Porcentaje"
         Object.Width           =   1482
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Categorías"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1020
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   6855
      Begin AACombo99.AACombo cPlazo 
         Height          =   315
         Left            =   840
         TabIndex        =   5
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
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
         Text            =   ""
      End
      Begin AACombo99.AACombo cCliente 
         Height          =   315
         Left            =   4440
         TabIndex        =   3
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
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
         Text            =   ""
      End
      Begin AACombo99.AACombo cArticulo 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
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
         Text            =   ""
      End
      Begin VB.TextBox tPorcentaje 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   4440
         MaxLength       =   5
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "P&orcentaje:"
         Height          =   255
         Left            =   3480
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Plazo:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   735
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
         Caption         =   "&Cliente:"
         Height          =   255
         Left            =   3480
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   5445
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "sucursal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label labIngresados 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingresados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   6855
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCategDto.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCategDto.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCategDto.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCategDto.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCategDto.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCategDto.frx":0BA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCategDto.frx":0EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCategDto.frx":11D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCategDto.frx":14F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCategDto.frx":180C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCategDto.frx":1B26
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCategDto.frx":1E40
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
      Begin VB.Menu MnuLinea1 
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
      Begin VB.Menu MnuLinea2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRefrescar 
         Caption         =   "&Refrescar"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "&Del Formulario"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuAcceso 
      Caption         =   "Acceso"
      Visible         =   0   'False
      Begin VB.Menu MnuListaArticulos 
         Caption         =   "Lista de Articulos"
      End
   End
End
Attribute VB_Name = "frmCatDto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim itmx As ListItem

'Booleanas.---------------------------------------------
Private sNuevoDto, sModificoDto As Boolean

'Long.------------------------------------------------------
Private CodCatArticulo As Long
Private CodCatCliente As Long
Private CodCatPlazo As Long

'Resultset.----------------------------------------------
Private RsAuxDto As rdoResultset

Private Sub cArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cArticulo.ListIndex > -1 Then cCliente.SetFocus
End Sub

Private Sub cArticulo_LostFocus()
    cArticulo.SelStart = 0
End Sub
Private Sub cPlazo_GotFocus()
    With cPlazo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub cPlazo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cPlazo.ListIndex > -1 Then Foco tPorcentaje
End Sub

Private Sub cPlazo_LostFocus()
    cPlazo.SelStart = 0
End Sub

Private Sub Form_Activate()
On Error Resume Next

    Screen.MousePointer = vbDefault
    DoEvents
        
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    
    ObtengoSeteoForm Me, 1500, 1000, 7245, 6390
    SetearLView lvValores.FullRow Or lvValores.Grilla, lvIngresados
    Status.Panels(5).Text = "Refrescar [F5]"

    'Inicializo booleanas.-----------------------------
    sNuevoDto = False
    sModificoDto = False
    
    'Oculto campos del form.-----------------------
    DeshabilitoCampos
    
    'Cargo Combos auxiliares.-------------------------
    CargoTipoPlazos
    CargoCatClientes
    CargoCatArticulos
    
    BuscoDescuentosAsignados
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrio un error inesperado al iniciar el formulario.", Err.Description
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If cCliente.ListIndex = -1 And cArticulo.ListIndex = -1 And cPlazo.ListIndex = -1 And _
        Trim(tPorcentaje.Text) = "" Then Exit Sub
            
    
    If sNuevoDto Or sModificoDto Then
    

        If MsgBox("Se han modificado datos en la ficha." & Chr(13) _
            & "¿Confirma almacenar los datos ingresados.?", vbQuestion + vbYesNo, "GRABAR " & Me.Caption) = vbYes Then
            
            If cCliente.ListIndex = -1 And cArticulo.ListIndex = -1 And cPlazo.ListIndex = -1 And _
                Not IsNumeric(tPorcentaje.Text) Then
                
                MsgBox "Los datos ingresados no son correctos.", vbExclamation, "ATENCIÓN"
                Cancel = True
                Exit Sub
            End If
            
            If sNuevoDto Then
                GrabarNuevoDescuento
                If sNuevoDto Then Cancel = True
            Else
                ModificoDescuento
                If sModificoDto Then Cancel = True
            End If
        End If
    End If

End Sub

Private Sub Form_Resize()
On Error Resume Next
    lvIngresados.Width = Me.ScaleWidth - (lvIngresados.Left * 2)
    lvIngresados.Height = Me.ScaleHeight - (lvIngresados.Top + Status.Height + 70)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    GuardoSeteoForm Me
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
End Sub

Private Sub Label1_Click()
    Foco cArticulo
End Sub

Private Sub Label2_Click()
    Foco cCliente
End Sub

Private Sub Label3_Click()
    Foco tPorcentaje
End Sub

Private Sub lvIngresados_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)

    lvIngresados.SortKey = ColumnHeader.Index - 1
    lvIngresados.Sorted = True

End Sub

Private Sub lvIngresados_ItemClick(ByVal Item As ComctlLib.ListItem)

    Botones True, True, True, False, False, Toolbar1, Me

End Sub

Private Sub Label5_Click()
    Foco cPlazo
End Sub

Private Sub lvIngresados_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lvIngresados.SelectedItem Is Nothing Then Exit Sub
    lvIngresados.HitTest x, y
    If Button = 2 And lvIngresados.ListItems.Count > 0 Then
        PopupMenu MnuAcceso
    End If
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

Private Sub MnuListaArticulos_Click()
On Error GoTo ErrMLA
    If lvIngresados.SelectedItem Is Nothing Then Exit Sub
    Dim liAyuda As New clsListadeAyuda
    Cons = "Select ArtID, Artículo = ArtNombre From Articulo, ArticuloFacturacion " _
        & " Where AFaCategoriaD = " & Mid(lvIngresados.SelectedItem.Key, 1, InStr(1, lvIngresados.SelectedItem.Key, ":") - 1) _
        & " And ArtID = AfaArticulo"
    liAyuda.ActivoListaAyuda Cons, False, miConexion.TextoConexion(logComercio), 5000
    Set liAyuda = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrMLA:
    clsGeneral.OcurrioError "Ocurrio un error al acceder a la lista de artículos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuModificar_Click()
    AccionModificar
End Sub
Private Sub MnuNuevo_Click()
    AccionNuevo
End Sub

Private Sub MnuRefrescar_Click()
    If Not sNuevoDto And Not sModificoDto Then AccionRefrescar
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Sub AccionNuevo()

    'Prendo Señal que es uno nuevo.---------------------
    sNuevoDto = True

    'Habilito y Desabilito Botones.--------------------------
    Botones False, False, False, True, True, Toolbar1, Me

    'Limpio y Muestro controles.---------------------------
    HabilitoCampos
    cArticulo.SetFocus

End Sub

Sub AccionGrabar()

    If cCliente.ListIndex > -1 And cArticulo.ListIndex > -1 And cPlazo.ListIndex > -1 And _
         IsNumeric(tPorcentaje.Text) Then
        
        If MsgBox("¿Confirma almacenar los datos ingresados.?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then
            If sNuevoDto Then
                GrabarNuevoDescuento
            Else
                ModificoDescuento
            End If
        End If
    Else
        MsgBox "Los datos ingresados no son correctos.", vbExclamation, "ATENCIÓN"
    End If

End Sub

Sub AccionEliminar()
    
    If MsgBox("¿Confirma eliminar la categoría de descuento seleccionada?", vbQuestion + vbYesNo, "ELIMINAR") = vbYes Then
        Screen.MousePointer = vbHourglass
        On Error GoTo ErrAE
        
        Cons = "Select * From CategoriaDescuento" _
                & " Where CDtCatArticulo = " & Mid(lvIngresados.SelectedItem.Key, 1, InStr(1, lvIngresados.SelectedItem.Key, ":") - 1) _
                & " And CDtCatCliente = " & Mid(lvIngresados.SelectedItem.Key, InStr(1, lvIngresados.SelectedItem.Key, ":") + 1, InStr(1, lvIngresados.SelectedItem.Key, "-") - InStr(1, lvIngresados.SelectedItem.Key, ":") - 1) _
                & " And CDtCatPlazo = " & Right(lvIngresados.SelectedItem.Key, Len(lvIngresados.SelectedItem.Key) - InStr(1, lvIngresados.SelectedItem.Key, "-"))
        
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If RsAux.EOF Then
            MsgBox "Otra terminal pudo eliminar el descuento seleccionado, verifique.", vbExclamation, "ATENCIÓN"
        Else
            RsAux.Delete
        End If
        RsAux.Close
        BuscoDescuentosAsignados
        Screen.MousePointer = vbDefault
    End If
   Exit Sub
 
ErrAE:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrio un error al intentar eliminar el descuento seleccionado.", Err.Description
    
End Sub

Sub AccionCancelar()

    'Apago señales.-------------------------------------------
    sNuevoDto = False
    sModificoDto = False
    'Limpio y oculto controles.---------------------------
    DeshabilitoCampos
    BuscoDescuentosAsignados
        
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        
        Case "nuevo"
            AccionNuevo
            
        Case "modificar"
            AccionModificar
        
        Case "eliminar"
            AccionEliminar
        
        Case "grabar"
            AccionGrabar
        
        Case "cancelar"
            AccionCancelar
            
        Case "refrescar"
            If Not sNuevoDto And Not sModificoDto Then AccionRefrescar
        
        Case "salir"
            Unload Me
            
    End Select

End Sub

Private Sub BuscoDescuentosAsignados()
On Error GoTo ErrBDA
    
    lvIngresados.ListItems.Clear
    labIngresados.Caption = " Descuentos Ingresados:"
    
    Cons = "Select CategoriaDescuento.*, CArNombre, CClNombre, TCuAbreviacion From CategoriaDescuento, CategoriaArticulo, CategoriaCliente, TipoCuota" _
            & " Where CDtCatArticulo = CArCodigo And CDtCatCliente = CClCodigo " _
            & " And CDtCatPlazo = TCuCodigo" _
            & " Order by CDtCatArticulo, CDtCatCliente, CDtCatPlazo"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    'Seteo los botones.-------------------------------------------------
    Botones True, False, False, False, False, Toolbar1, Me
    
    I = 0
    Do While Not RsAux.EOF
        Set itmx = lvIngresados.ListItems.Add(, RsAux!CDtCatArticulo & ":" & RsAux!CDtCatCliente & "-" & RsAux!CDtCatPlazo, Trim(RsAux!CArNombre))
        itmx.SubItems(1) = Trim(RsAux!CClNombre)
        itmx.SubItems(2) = Trim(RsAux!TCuAbreviacion)
        itmx.SubItems(3) = Format(RsAux!CDtPorcentaje, "#,##0.00")
        RsAux.MoveNext
        I = I + 1
    Loop
    RsAux.Close
    labIngresados.Caption = " Descuentos Ingresados: " & I
    Screen.MousePointer = 0
    Exit Sub
    
ErrBDA:
    clsGeneral.OcurrioError "Ocurrio un error al cargar los descuentos ingresados.", Err.Description
    Screen.MousePointer = vbDefault
End Sub

Private Sub CargoTipoPlazos()
On Error GoTo ErrCC

    Cons = "Select TCuCodigo, TCuAbreviacion From TipoCuota Order by TCuAbreviacion"
    CargoCombo Cons, cPlazo, ""
    Exit Sub

ErrCC:
    clsGeneral.OcurrioError "Ocurrio un error al cargar los tipos de Descuentos."
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub CargoCatClientes()
On Error GoTo ErrCC

    Cons = "Select CClCodigo, CClNombre From CategoriaCliente Order by CClNombre"
    CargoCombo Cons, cCliente, ""
    Exit Sub

ErrCC:
    clsGeneral.OcurrioError "Ocurrio un error al cargar las categorías de clientes."
    Screen.MousePointer = vbDefault

End Sub

Private Sub CargoCatArticulos()
On Error GoTo ErrCC

    Cons = "Select CArCodigo, CArNombre From CategoriaArticulo Order by CArNombre"
    CargoCombo Cons, cArticulo, ""
    Exit Sub

ErrCC:
    clsGeneral.OcurrioError "Ocurrio un error al cargar las categorías de Artículos."
    Screen.MousePointer = vbDefault

End Sub

Private Sub DeshabilitoCampos()

    cCliente.BackColor = Inactivo
    cCliente.Enabled = False
    cCliente.ListIndex = -1
    tPorcentaje.BackColor = Inactivo
    tPorcentaje.Enabled = False
    tPorcentaje.Text = ""
    cArticulo.Enabled = False
    cArticulo.BackColor = Inactivo
    cArticulo.ListIndex = -1
    cPlazo.Enabled = False
    cPlazo.BackColor = Inactivo
    cPlazo.ListIndex = -1
    lvIngresados.Enabled = True
    
End Sub

Private Sub HabilitoCampos()

    cCliente.BackColor = Obligatorio
    cCliente.Enabled = True
    tPorcentaje.BackColor = Obligatorio
    tPorcentaje.Enabled = True
    cArticulo.Enabled = True
    cArticulo.BackColor = Obligatorio
    cPlazo.Enabled = True
    cPlazo.BackColor = Obligatorio
    lvIngresados.Enabled = False
    
End Sub

Private Sub GrabarNuevoDescuento()

    Screen.MousePointer = vbHourglass
    On Error GoTo ErrGNE
    
    Cons = "Select * From CategoriaDescuento" _
        & " Where CDtCatArticulo = " & cArticulo.ItemData(cArticulo.ListIndex) _
        & " And CDtCatCliente = " & cCliente.ItemData(cCliente.ListIndex) _
        & " And CDtCatPlazo = " & cPlazo.ItemData(cPlazo.ListIndex)
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        Screen.MousePointer = vbDefault
        MsgBox "El Descuento ya fue ingresado.", vbExclamation, "ATENCION"
   Else
        RsAux.AddNew
        GuardoDatosEnResultset
        RsAux.Update
    End If
    AccionCancelar
    Screen.MousePointer = vbDefault
    Exit Sub
 
ErrGNE:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrio un error al intentar almacenar la información.", Err.Description
End Sub

Private Sub AccionRefrescar()
On Error GoTo ErrAR
        
    Screen.MousePointer = vbHourglass
    'Cargo Combos auxiliares.-------------------------
    Cons = "Select TCuCodigo, TCuAbreviacion From TipoCuota Order by TCuAbreviacion"
    CargoCombo Cons, cPlazo, cPlazo.Text
    
    Cons = "Select CArCodigo, CArNombre From CategoriaArticulo Order by CArNombre"
    CargoCombo Cons, cArticulo, cArticulo.Text
    
    Cons = "Select CClCodigo, CClNombre From CategoriaCliente Order by CClNombre"
    CargoCombo Cons, cCliente, cCliente.Text
    
    BuscoDescuentosAsignados
    
    Screen.MousePointer = vbDefault
    Exit Sub

ErrAR:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error inesperado al refrescar la información."
    
End Sub

Private Sub cArticulo_GotFocus()
    cArticulo.SelStart = 0
    cArticulo.SelLength = Len(cArticulo.Text)
End Sub

Private Sub cCliente_GotFocus()
    cCliente.SelStart = 0
    cCliente.SelLength = Len(cCliente.Text)
End Sub

Private Sub cCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cCliente.ListIndex > -1 Then Foco cPlazo
End Sub

Private Sub cCliente_LostFocus()
    cCliente.SelLength = 0
End Sub

Private Sub tPorcentaje_GotFocus()
On Error Resume Next
    Foco tPorcentaje
End Sub

Private Sub ModificoDescuento()
On Error GoTo ErrMD

    Cons = "Select * From CategoriaDescuento" _
            & " Where CDtCatArticulo = " & cArticulo.ItemData(cArticulo.ListIndex) _
            & " And CDtCatCliente = " & cCliente.ItemData(cCliente.ListIndex) _
            & " And CDtCatPlazo = " & cPlazo.ItemData(cPlazo.ListIndex)
    
    Set RsAuxDto = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAuxDto.EOF Then
        If Not (CCur(lvIngresados.SelectedItem.SubItems(3)) = RsAuxDto!CDtPorcentaje _
            And RsAuxDto!CDtCatArticulo = Mid(lvIngresados.SelectedItem.Key, 1, InStr(1, lvIngresados.SelectedItem.Key, ":") - 1) _
            And Mid(lvIngresados.SelectedItem.Key, InStr(1, lvIngresados.SelectedItem.Key, ":") + 1, InStr(1, lvIngresados.SelectedItem.Key, "-") - InStr(1, lvIngresados.SelectedItem.Key, ":") - 1) = RsAuxDto!CDtCatCliente _
            And Right(lvIngresados.SelectedItem.Key, Len(lvIngresados.SelectedItem.Key) - InStr(1, lvIngresados.SelectedItem.Key, "-")) = RsAuxDto!CDtCatPlazo) Then
        
        'Modifico algo.--------------------------------------------------
        
            If Mid(lvIngresados.SelectedItem.Key, 1, InStr(1, lvIngresados.SelectedItem.Key, ":") - 1) <> RsAuxDto!CDtCatArticulo _
                Or Mid(lvIngresados.SelectedItem.Key, InStr(1, lvIngresados.SelectedItem.Key, ":") + 1, InStr(1, lvIngresados.SelectedItem.Key, "-") - InStr(1, lvIngresados.SelectedItem.Key, ":") - 1) <> RsAuxDto!CDtCatCliente _
                Or Right(lvIngresados.SelectedItem.Key, Len(lvIngresados.SelectedItem.Key) - InStr(1, lvIngresados.SelectedItem.Key, "-")) <> RsAuxDto!CDtCatPlazo Then
                
                'Ya existe uno así.--------------------------------------------
                
                RsAuxDto.Close
                MsgBox "Ya existe un descuento con esas características, verifique.", vbExclamation, "ATENCIÓN"
                AccionCancelar
                Exit Sub
            End If
        End If
    End If
    RsAuxDto.Close
    
    Cons = "Select * From CategoriaDescuento" _
        & " Where CDtCatArticulo = " & Mid(lvIngresados.SelectedItem.Key, 1, InStr(1, lvIngresados.SelectedItem.Key, ":") - 1) _
        & " And CDtCatCliente = " & Mid(lvIngresados.SelectedItem.Key, InStr(1, lvIngresados.SelectedItem.Key, ":") + 1, InStr(1, lvIngresados.SelectedItem.Key, "-") - InStr(1, lvIngresados.SelectedItem.Key, ":") - 1) _
        & " And CDtCatPlazo = " & Right(lvIngresados.SelectedItem.Key, Len(lvIngresados.SelectedItem.Key) - InStr(1, lvIngresados.SelectedItem.Key, "-"))

    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        MsgBox "Otra terminal pudo eliminar la categoría, verfique.", vbExclamation, "ATENCIÓN"
    Else
        RsAux.Edit
        GuardoDatosEnResultset
        RsAux.Update
    End If
    RsAux.Close
    AccionCancelar
    Exit Sub

ErrMD:
    clsGeneral.OcurrioError "Ocurrio un error al modificar los datos.", Err.Description
    
End Sub

Private Sub GuardoDatosEnResultset()
        
    RsAux!CDtCatArticulo = cArticulo.ItemData(cArticulo.ListIndex)
    RsAux!CDtCatCliente = cCliente.ItemData(cCliente.ListIndex)
    RsAux!CDtCatPlazo = cPlazo.ItemData(cPlazo.ListIndex)
    RsAux!CDtPorcentaje = tPorcentaje.Text

End Sub

Private Sub tPorcentaje_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then AccionGrabar

End Sub

Private Sub AccionModificar()
    
    Cons = "Select * From CategoriaDescuento" _
        & " Where CDtCatArticulo = " & Mid(lvIngresados.SelectedItem.Key, 1, InStr(1, lvIngresados.SelectedItem.Key, ":") - 1) _
        & " And CDtCatCliente = " & Mid(lvIngresados.SelectedItem.Key, InStr(1, lvIngresados.SelectedItem.Key, ":") + 1, InStr(1, lvIngresados.SelectedItem.Key, "-") - InStr(1, lvIngresados.SelectedItem.Key, ":") - 1) _
        & " And CDtCatPlazo = " & Right(lvIngresados.SelectedItem.Key, Len(lvIngresados.SelectedItem.Key) - InStr(1, lvIngresados.SelectedItem.Key, "-"))

    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "El descuento seleccionado ha sido eliminado, verifique.", vbExclamation, "ATENCIÓN"
        BuscoDescuentosAsignados
    Else
        HabilitoCampos
        BuscoCodigoEnCombo cArticulo, RsAux!CDtCatArticulo
        BuscoCodigoEnCombo cCliente, RsAux!CDtCatCliente
        BuscoCodigoEnCombo cPlazo, RsAux!CDtCatPlazo
        tPorcentaje.Text = RsAux!CDtPorcentaje
        RsAux.Close
        sModificoDto = True
        Botones False, False, False, True, True, Toolbar1, Me
        cArticulo.SetFocus
    End If

End Sub
