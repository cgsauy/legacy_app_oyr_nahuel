VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form frmFiltros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtros de Consulta"
   ClientHeight    =   3360
   ClientLeft      =   5100
   ClientTop       =   2520
   ClientWidth     =   4080
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFiltros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4080
   Begin VB.TextBox tVigencia 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox tPrecio 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   2100
      Width           =   1275
   End
   Begin VB.TextBox tLista 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1320
      Width           =   3135
   End
   Begin VB.CommandButton bAplicar 
      Caption         =   "&Aplicar"
      Height          =   315
      Left            =   3180
      TabIndex        =   18
      Top             =   2760
      Width           =   795
   End
   Begin VB.CheckBox cHabilitado 
      Alignment       =   1  'Right Justify
      Caption         =   "&Habilitado p/venta"
      Height          =   195
      Left            =   2280
      TabIndex        =   17
      Top             =   2100
      Width           =   1695
   End
   Begin VB.CheckBox cExclusivo 
      Alignment       =   1  'Right Justify
      Caption         =   "Art. &Exclusivo"
      Height          =   195
      Left            =   2280
      TabIndex        =   16
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox tTipo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   420
      Width           =   2475
   End
   Begin VB.TextBox tMarca 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   720
      Width           =   2475
   End
   Begin VB.TextBox tProveedor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1020
      Width           =   3135
   End
   Begin VB.TextBox tGrupo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   2475
   End
   Begin AACombo99.AACombo cPlan 
      Height          =   315
      Left            =   840
      TabIndex        =   11
      Top             =   1740
      Width           =   795
      _ExtentX        =   1402
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
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   3105
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   6694
            Key             =   "help"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "&Vigencia:"
      Height          =   195
      Left            =   60
      TabIndex        =   14
      Top             =   2415
      Width           =   675
   End
   Begin VB.Label Label6 
      Caption         =   "P&recio:"
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   2115
      Width           =   555
   End
   Begin VB.Label lLista 
      BackStyle       =   0  'Transparent
      Caption         =   "&Lista ="
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   1335
      Width           =   795
   End
   Begin VB.Label lPlan 
      Caption         =   "&Plan ="
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   1815
      Width           =   795
   End
   Begin VB.Label lTipo 
      BackStyle       =   0  'Transparent
      Caption         =   "&Tipo ="
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   435
      Width           =   795
   End
   Begin VB.Label lMarca 
      BackStyle       =   0  'Transparent
      Caption         =   "&Marca ="
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   735
      Width           =   795
   End
   Begin VB.Label lProveedor 
      BackStyle       =   0  'Transparent
      Caption         =   "&Prov ="
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   1035
      Width           =   795
   End
   Begin VB.Label lGrupo 
      BackStyle       =   0  'Transparent
      Caption         =   "&Grupo ="
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   135
      Width           =   795
   End
End
Attribute VB_Name = "frmFiltros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public prmOK As Boolean

Public prmGrupo As String
Public prmTipo As String
Public prmMarca As String
Public prmProveedor As String
Public prmLista As String
Public prmPlan As String
Public prmExclusivo As String
Public prmHabilitado As String
Public prmPrecio As String
Public prmVigencia As String

Private Enum enuControl
    cGrupo = 1
    cTipo = 2
    cMarca = 3
    cProveedor = 4
    cLista = 5
End Enum
    

Private Sub bAplicar_Click()
    On Error Resume Next
    
    With tPrecio
        If Trim(.Text) <> "" Then
            If Not ((Mid(.Text, 1, 1) = "<" Or Mid(.Text, 1, 1) = ">") And IsNumeric(Mid(.Text, 2))) Then
                MsgBox "Los datos para el filtro 'Precio' no son correctos." & vbCrLf & vbCrLf & _
                            "Formatos: >999 ó <999", vbExclamation, "Filtro Incorrecto"
                Foco tPrecio: Exit Sub
            End If
        End If
    End With
    
    If Trim(tVigencia.Text) <> "" Then
        If ValidoPeriodoFechas(tVigencia.Text) = "" Then
            MsgBox "Los datos para el filtro 'Vigencia del Precio' no son correctos." & vbCrLf & vbCrLf & _
                            "Formatos: >... , <... ó E...Y...", vbExclamation, "Filtro Incorrecto"
            Foco tVigencia: Exit Sub
        End If
    End If
    
    '1- True     0- False
    With tGrupo
        If Val(.Tag) <> 0 Then
            If .BackColor = vbWindowBackground Then prmGrupo = 1 Else prmGrupo = 0
            prmGrupo = prmGrupo & Val(.Tag)
        End If
    End With
    
    With tTipo
        If Val(.Tag) <> 0 Then
            If .BackColor = vbWindowBackground Then prmTipo = 1 Else prmTipo = 0
            prmTipo = prmTipo & Val(.Tag)
        End If
    End With
    
    With tMarca
        If Val(.Tag) <> 0 Then
            If .BackColor = vbWindowBackground Then prmMarca = 1 Else prmMarca = 0
            prmMarca = prmMarca & Val(.Tag)
        End If
    End With
    
    With tProveedor
        If Val(.Tag) <> 0 Then
            If .BackColor = vbWindowBackground Then prmProveedor = 1 Else prmProveedor = 0
            prmProveedor = prmProveedor & Val(.Tag)
        End If
    End With
    
    With tLista
        If Val(.Tag) <> 0 Then
            If .BackColor = vbWindowBackground Then prmLista = 1 Else prmLista = 0
            prmLista = prmLista & Val(.Tag)
        End If
    End With
    
    With cPlan
        If .ListIndex <> -1 Then
            If .BackColor = vbWindowBackground Then prmPlan = 1 Else prmPlan = 0
            prmPlan = prmPlan & .ItemData(.ListIndex)
        End If
    End With
    
    With cExclusivo
        If .Value <> vbGrayed Then
            If .Value = vbChecked Then prmExclusivo = 1 Else prmExclusivo = 0
        End If
    End With
    
    With cHabilitado
        If .Value <> vbGrayed Then
            If .Value = vbChecked Then prmHabilitado = 1 Else prmHabilitado = 0
        End If
    End With
    
    With tPrecio
        If Trim(.Text) <> "" Then
            If (Mid(.Text, 1, 1) = "<" Or Mid(.Text, 1, 1) = ">") And IsNumeric(Mid(.Text, 2)) Then
                prmPrecio = Trim(.Text)
            End If
        End If
    End With
    
    With tVigencia
        If Trim(.Text) <> "" Then
            prmVigencia = RetornoFormatoFechaConsulta(.Text)
        End If
    End With
    
    prmOK = True
    Unload Me
End Sub

Private Sub bAplicar_GotFocus()
    Status.Panels("help").Text = ""
End Sub

Private Sub cExclusivo_GotFocus()
    Status.Panels("help").Text = ""
End Sub

Private Sub cExclusivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cHabilitado.SetFocus
End Sub

Private Sub cExclusivo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then cExclusivo.Value = vbGrayed
End Sub

Private Sub cHabilitado_GotFocus()
    Status.Panels("help").Text = ""
End Sub

Private Sub cHabilitado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bAplicar.SetFocus
End Sub

Private Sub cHabilitado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then cHabilitado.Value = vbGrayed
End Sub

Private Sub cPlan_GotFocus()
    Status.Panels("help").Text = "[F3]- Cambia modo de filtrado."
End Sub

Private Sub cPlan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then StateControl cPlan, lPlan
End Sub

Private Sub cPlan_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tPrecio
End Sub

Private Sub Form_Load()
On Error Resume Next
    CentroForm Me
    InicializoForm
End Sub

Private Sub InicializoForm()
On Error Resume Next
    prmOK = False
    cons = "Select PlaCodigo, PlaNombre from TipoPlan Order by PlaNombre"
    CargoCombo cons, cPlan
    
    tGrupo.Text = ""
    tTipo.Text = ""
    tMarca.Text = ""
    tProveedor.Text = ""
    tLista.Text = ""
    tVigencia.Text = ">" & Format(DateAdd("yyyy", -1, Now), "dd/mm/yyyy")
    
    cPlan.Text = ""
    tPrecio.Text = ""
    cHabilitado.Value = vbGrayed
    cExclusivo.Value = vbGrayed
        
    prmGrupo = ""
    prmTipo = ""
    prmMarca = ""
    prmProveedor = ""
    prmLista = ""
    prmPlan = ""
    prmExclusivo = ""
    prmHabilitado = ""
    prmPrecio = ""
    prmVigencia = ""

    lGrupo.Tag = "&Grupo"
    lTipo.Tag = "&Tipo"
    lMarca.Tag = "&Marca"
    lProveedor.Tag = "&Prov"
    lLista.Tag = "&Lista"
    lPlan.Tag = "&Plan"
    
End Sub

Private Sub tGrupo_Change()
    tGrupo.Tag = 0
End Sub

Private Sub tGrupo_GotFocus()
    Status.Panels("help").Text = "[F3]- Cambia modo de filtrado."
End Sub

Private Sub tGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then StateControl tGrupo, lGrupo
End Sub

Private Sub tGrupo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tGrupo.Text) = "" Then Foco tTipo: Exit Sub
        If Val(tGrupo.Tag) <> 0 Then Foco tTipo: Exit Sub
                        
        AyudaControl cGrupo, tGrupo
    End If
    
End Sub

Private Function AyudaControl(Tipo As enuControl, mControl As TextBox)

    On Error GoTo errBuscaG
    Screen.MousePointer = 11
    
    Dim aTexto As String
    Dim aQ As Integer, aIDSel As Long
    aQ = 0: aIDSel = 0
    mControl.Text = Replace(Trim(mControl.Text), " ", "%")
    Select Case Tipo
        Case enuControl.cGrupo     'Grupo
                    cons = "Select GruCodigo as Codigo, GruNombre  as Nombre from Grupo " & _
                               "Where GruNombre like '" & Trim(mControl.Text) & "%'" & _
                               " Order by Nombre"
    
        Case enuControl.cTipo      'Tipo
                    cons = "Select TipCodigo as Codigo, TipNombre  as Nombre from Tipo " & _
                               "Where TipNombre like '" & Trim(mControl.Text) & "%'" & _
                               " Order by Nombre"
        
        Case enuControl.cMarca      'Marca
                    cons = "Select MarCodigo as Codigo, MarNombre  as Nombre from Marca " & _
                               "Where MarNombre like '" & Trim(mControl.Text) & "%'" & _
                               " Order by Nombre"
        
        Case enuControl.cProveedor  'Proveedor
                    cons = "Select PMeCodigo as Codigo, PMeFantasia as Nombre From ProveedorMercaderia " & _
                               "Where PMeFantasia like '" & Trim(mControl.Text) & "%'" & _
                               "OR PMeNombre like '" & Trim(mControl.Text) & "%'" & _
                               " Order by Nombre"
                               
        Case enuControl.cLista  'Proveedor
                    cons = "Select LDPCodigo as Codigo, LDPDescripcion as Nombre From ListasdePrecios " & _
                               "Where LDPDescripcion like '" & Trim(mControl.Text) & "%'" & _
                               "OR LDPNombre like '" & Trim(mControl.Text) & "%'" & _
                               " Order by Nombre"
    End Select
        
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        aQ = 1
        aIDSel = rsAux!Codigo: aTexto = Trim(rsAux!Nombre)
        rsAux.MoveNext: If Not rsAux.EOF Then aQ = 2
    End If
    rsAux.Close
        
    Select Case aQ
        Case 0: MsgBox "No hay datos que coincidan con el texto ingersado.", vbExclamation, "No hay datos"
        
        Case 2:
                    Dim miLista As New clsListadeAyuda
                    aIDSel = miLista.ActivarAyuda(cBase, cons, 4000, 1, "Lista de Datos")
                    Me.Refresh
                    If aIDSel > 0 Then
                        aIDSel = miLista.RetornoDatoSeleccionado(0)
                        aTexto = miLista.RetornoDatoSeleccionado(1)
                    End If
                    Set miLista = Nothing
    End Select
        
    If aIDSel > 0 Then
        mControl.Text = Trim(aTexto)
        mControl.Tag = aIDSel
    End If
    Screen.MousePointer = 0
    Exit Function
    
errBuscaG:
    clsGeneral.OcurrioError "Error al buscar los datos.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub tLista_Change()
    tLista.Tag = 0
End Sub

Private Sub tLista_GotFocus()
    Status.Panels("help").Text = "[F3]- Cambia modo de filtrado."
End Sub

Private Sub tLista_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then StateControl tLista, lLista
End Sub

Private Sub tLista_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(tLista.Text) = "" Then Foco cPlan: Exit Sub
        If Val(tLista.Tag) <> 0 Then Foco cPlan: Exit Sub
                        
        AyudaControl cLista, tLista
    End If
End Sub

Private Sub tMarca_Change()
    tMarca.Tag = 0
End Sub

Private Sub tMarca_GotFocus()
    Status.Panels("help").Text = "[F3]- Cambia modo de filtrado."
End Sub

Private Sub tMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then StateControl tMarca, lMarca
End Sub

Private Sub tMarca_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(tMarca.Text) = "" Then Foco tProveedor: Exit Sub
        If Val(tMarca.Tag) <> 0 Then Foco tProveedor: Exit Sub
                        
        AyudaControl cMarca, tMarca
    End If
End Sub

Private Sub tPrecio_GotFocus()
    Status.Panels("help").Text = "Pude usar '>' y '<' para ingresar el valor del filtro."
End Sub

Private Sub tPrecio_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tVigencia
End Sub

Private Sub tProveedor_Change()
    tProveedor.Tag = 0
End Sub

Private Sub tProveedor_GotFocus()
    Status.Panels("help").Text = "[F3]- Cambia modo de filtrado."
End Sub

Private Sub tProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then StateControl tProveedor, lProveedor
End Sub

Private Sub tProveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(tProveedor.Text) = "" Then Foco tLista: Exit Sub
        If Val(tProveedor.Tag) <> 0 Then Foco tLista: Exit Sub
                        
        AyudaControl cProveedor, tProveedor
    End If
End Sub

Private Sub tTipo_Change()
    tTipo.Tag = 0
End Sub

Private Sub tTipo_GotFocus()
    Status.Panels("help").Text = "[F3]- Cambia modo de filtrado."
End Sub

Private Sub tTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then StateControl tTipo, lTipo
End Sub

Private Sub tTipo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Trim(tTipo.Text) = "" Then Foco tMarca: Exit Sub
        If Val(tTipo.Tag) <> 0 Then Foco tMarca: Exit Sub
                        
        AyudaControl cTipo, tTipo
    End If
    
End Sub

Private Function StateControl(mControl As Control, mCaption As Label)
    If mControl.BackColor = vbWindowBackground Then
        mControl.BackColor = vbInactiveBorder
        mCaption.Caption = Trim(mCaption.Tag) & " <>"
    Else
        mControl.BackColor = vbWindowBackground
        mCaption.Caption = Trim(mCaption.Tag) & " ="
    End If
End Function

Private Sub tVigencia_GotFocus()
    Status.Panels("help").Text = "Pude usar '>' , '<' ó 'E...Y...' para ingresar la fecha."
End Sub

Private Sub tVigencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cExclusivo.SetFocus
End Sub
