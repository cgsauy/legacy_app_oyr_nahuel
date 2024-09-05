VERSION 5.00
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form frmCondiciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Condiciones"
   ClientHeight    =   5745
   ClientLeft      =   2325
   ClientTop       =   2790
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCondiciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   7575
   Begin VB.TextBox txtTextoSiSi 
      Height          =   645
      Left            =   1500
      MaxLength       =   150
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   4000
      Width           =   5955
   End
   Begin VB.CheckBox chDefinitiva 
      Appearance      =   0  'Flat
      Caption         =   "De&finitiva"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4080
      TabIndex        =   9
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton bGrabarCerrar 
      Caption         =   "Grabar/Salir"
      Height          =   315
      Left            =   5160
      TabIndex        =   17
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton bEliminar 
      Caption         =   "Eliminar"
      Height          =   315
      Left            =   2760
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton bEvaluarTodo 
      Caption         =   "Evaluar c/c"
      Height          =   315
      Left            =   5280
      TabIndex        =   27
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox tExpresion 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   585
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   2940
      Width           =   7335
   End
   Begin AACombo99.AACombo cResSiSi 
      Height          =   315
      Left            =   1500
      TabIndex        =   8
      Top             =   3600
      Width           =   2475
      _ExtentX        =   4366
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
      Text            =   ""
   End
   Begin VB.CommandButton bCancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   6360
      TabIndex        =   18
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton bGrabar 
      Caption         =   "Grabar"
      Height          =   315
      Left            =   3960
      TabIndex        =   16
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox tprmSolicitud 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   21
      Text            =   "3200"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton bEvaluar 
      Caption         =   "&Evaluar"
      Height          =   315
      Left            =   6480
      TabIndex        =   6
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox tCondicion 
      Appearance      =   0  'Flat
      Height          =   825
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1740
      Width           =   7335
   End
   Begin VB.TextBox tDescripcion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      MaxLength       =   60
      TabIndex        =   3
      Top             =   1020
      Width           =   6375
   End
   Begin VB.TextBox tNombre 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   1
      Top             =   720
      Width           =   6375
   End
   Begin AACombo99.AACombo cCaminoSiSi 
      Height          =   315
      Left            =   1500
      TabIndex        =   13
      Top             =   4740
      Width           =   2475
      _ExtentX        =   4366
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
      Text            =   ""
   End
   Begin AACombo99.AACombo cCaminoSiNo 
      Height          =   315
      Left            =   5400
      TabIndex        =   15
      Top             =   4740
      Width           =   2055
      _ExtentX        =   3625
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
      Text            =   ""
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "&Texto Si ""Si"":"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   1395
   End
   Begin VB.Label lCondSiNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Camino Si ""&No"":"
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   4800
      Width           =   1155
   End
   Begin VB.Label lCondSISI 
      BackStyle       =   0  'Transparent
      Caption         =   "Camino Si ""&Si"":"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4770
      Width           =   1395
   End
   Begin VB.Label lCabezal2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      Height          =   45
      Left            =   60
      TabIndex        =   25
      Top             =   5160
      Width           =   4875
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingreso de la Condición"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   780
      TabIndex        =   24
      Top             =   120
      Width           =   2835
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "frmCondiciones.frx":030A
      Top             =   20
      Width           =   480
   End
   Begin VB.Label lCabezal 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      Height          =   45
      Left            =   0
      TabIndex        =   23
      Top             =   540
      Width           =   4875
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "[prmSolicitud]:"
      Height          =   255
      Left            =   5460
      TabIndex        =   22
      Top             =   165
      Width           =   1035
   End
   Begin VB.Label lResultado 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   20
      Top             =   2685
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Resultado:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2685
      Width           =   915
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Condición a Evaluar:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1500
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Resolución Si ""Si"":"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3630
      Width           =   1395
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Descripción:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1065
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Nombre:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   765
      Width           =   735
   End
End
Attribute VB_Name = "frmCondiciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public prmIdValor As Long
Public prmGrabo As Boolean

Dim rsEva As rdoResultset

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bEliminar_Click()
    AccionEliminar
End Sub

Private Sub bEvaluar_Click()
    AccionEvaluar
End Sub

Private Sub bEvaluarTodo_Click()
    AccionEvaluarTodo
End Sub

Private Sub bGrabar_Click()
    AccionGrabar
End Sub

Private Sub bGrabarCerrar_Click()
    AccionGrabar cExit:=True
End Sub

Private Sub cCaminoSiNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And cCaminoSiNo.ListIndex > -1 Then
        'Cargo condición
        prmIdValor = cCaminoSiNo.ItemData(cCaminoSiNo.ListIndex)
        CargoDatosValor
        tNombre.SetFocus
        bEliminar.Enabled = True
    End If
End Sub

Private Sub cCaminoSiNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bGrabar.SetFocus
End Sub

Private Sub cCaminoSiSi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And cCaminoSiSi.ListIndex > -1 Then
        'Cargo condición
        prmIdValor = cCaminoSiSi.ItemData(cCaminoSiSi.ListIndex)
        CargoDatosValor
        tNombre.SetFocus
        bEliminar.Enabled = True
    End If
End Sub

Private Sub cCaminoSiSi_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cCaminoSiNo.SetFocus
End Sub


Private Sub chDefinitiva_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cResSiSi.ListIndex > -1 Then
            txtTextoSiSi.SetFocus
        Else
            cCaminoSiSi.SetFocus
        End If
    End If
End Sub

Private Sub cResSiSi_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then chDefinitiva.SetFocus
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    'Center form
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    prmGrabo = False
    
    InicializoControles
    If prmIdValor <> 0 Then
        CargoDatosValor
        bEliminar.Enabled = True
    End If
    Screen.MousePointer = 0
    
    Exit Sub
errLoad:
    clsGeneral.OcurrioError "Error al inicializar el formulario", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionEvaluar()
    
    On Error GoTo errEvaluar
    lResultado.Caption = ""
    Screen.MousePointer = 11
    
    Dim mOBJ As New clsResolAuto
    Dim aResult As Variant, aExp As String
    
    aResult = mOBJ.EvaluoCondicion(cBase, Val(tprmSolicitud.Text), Trim(tCondicion.Text), aExp)
    Set mOBJ = Nothing
    
    lResultado.Caption = CBool(aResult)
    tExpresion.Text = aExp
    
    Screen.MousePointer = 0
    Exit Sub
    
errEvaluar:
    clsGeneral.OcurrioError "Error al evaluar la condición", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionEvaluarTodo()
    On Error GoTo errEvaluar
    
    AccionGrabar cPregunta:=False
    If prmIdValor = 0 Then Exit Sub
    
    Dim aResult As Variant
    
    lResultado.Caption = ""
    Screen.MousePointer = 11
    
    Dim mOBJ As New clsResolAuto
    aResult = mOBJ.VerCaminoAlResolver(cBase, Val(tprmSolicitud.Text), idCondInicial:=prmIdValor)
    Set mOBJ = Nothing
    
    lResultado.Caption = CBool(aResult)
    
    Screen.MousePointer = 0
    Exit Sub
    
errEvaluar:
    clsGeneral.OcurrioError "Error al evaluar la condición", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionGrabar(Optional cPregunta As Boolean = True, Optional cExit As Boolean = False)

    If Not ValidoCampos Then Exit Sub
    If cPregunta Then
        If MsgBox("Confirma grabar los datos para la condición: " & Trim(tNombre.Text), vbQuestion + vbYesNo, "Grabar") = vbNo Then Exit Sub
    End If
    
    On Error GoTo errGrabar
    Screen.MousePointer = 11
    
    Cons = "Select * from ValoresCalculados Where VCaCodigo = " & prmIdValor
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then RsAux.Edit Else RsAux.AddNew
    
    RsAux!VCaNombre = Trim(tNombre.Text)
    RsAux!VCaTipo = 2
    If Trim(tDescripcion.Text) <> "" Then RsAux!VCaDescripcion = Trim(tDescripcion.Text) Else RsAux!VCaDescripcion = Null
    RsAux!VCaTexto = Trim(tCondicion.Text)
    
    If cResSiSi.ListIndex <> -1 Then RsAux!VCaResolucionSiSi = cResSiSi.ItemData(cResSiSi.ListIndex) Else RsAux!VCaResolucionSiSi = Null
    If cCaminoSiSi.ListIndex <> -1 Then RsAux!VCaCaminoSiSi = cCaminoSiSi.ItemData(cCaminoSiSi.ListIndex) Else RsAux!VCaCaminoSiSi = Null
    If cCaminoSiNo.ListIndex <> -1 Then RsAux!VCaCaminoSiNo = cCaminoSiNo.ItemData(cCaminoSiNo.ListIndex) Else RsAux!VCaCaminoSiNo = Null
    RsAux("VCaResDefinitiva") = IIf(chDefinitiva.Value = 1, True, False)
    
    If Trim(txtTextoSiSi.Text) <> "" Then
        RsAux("VCaTextoResDef") = Replace(txtTextoSiSi.Text, vbCrLf, " ")
    Else
        RsAux("VCaTextoResDef") = Null
    End If
    
    RsAux.Update: RsAux.Close
    
    prmGrabo = True
    
    If prmIdValor = 0 Then
        Cons = "Select * from ValoresCalculados Where VCaNombre = '" & Trim(tNombre.Text) & "'"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            prmIdValor = RsAux!VCaCodigo
            bEliminar.Enabled = True
        End If
        RsAux.Close
    End If
    
    Screen.MousePointer = 0
    
    If cExit Then Unload Me
    
    Exit Sub
errGrabar:
    clsGeneral.OcurrioError "Error al grabar la condición.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoDatosValor()

    On Error GoTo errCargar
    Screen.MousePointer = 11
    
    tNombre.Text = ""
    tDescripcion.Text = ""
    tCondicion.Text = ""
    tExpresion.Text = ""
    
        
    cResSiSi.Text = ""
    txtTextoSiSi.Text = ""
    cCaminoSiSi.Text = ""
    cCaminoSiNo.Text = ""
    chDefinitiva.Value = 0
    
    Cons = "Select * from ValoresCalculados Where VCaCodigo = " & prmIdValor
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
    
        tNombre.Text = Trim(RsAux!VCaNombre)
        If Not IsNull(RsAux!VCaDescripcion) Then tDescripcion.Text = Trim(RsAux!VCaDescripcion) Else tDescripcion.Text = ""
        tCondicion.Text = Trim(RsAux!VCaTexto)
        
        If Not IsNull(RsAux!VCaResolucionSiSi) Then BuscoCodigoEnCombo cResSiSi, RsAux!VCaResolucionSiSi
        If Not IsNull(RsAux!VCaCaminoSiSi) Then BuscoCodigoEnCombo cCaminoSiSi, RsAux!VCaCaminoSiSi
        If Not IsNull(RsAux!VCaCaminoSiNo) Then BuscoCodigoEnCombo cCaminoSiNo, RsAux!VCaCaminoSiNo
        If Not IsNull(RsAux("VCaResDefinitiva")) Then chDefinitiva.Value = IIf(RsAux("VCaResDefinitiva"), 1, 0)
        If Not IsNull(RsAux("VCaTextoResDef")) Then txtTextoSiSi.Text = Trim(RsAux("VCaTextoResDef"))
        
        
    End If
    
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Error al cargar los datos en la ficha.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lCabezal.Left = Me.ScaleLeft + 60: lCabezal.Width = Me.ScaleWidth - (lCabezal.Left * 2)
    lCabezal2.Left = lCabezal.Left: lCabezal2.Width = lCabezal.Width
    
End Sub

Private Function ValidoCampos() As Boolean

    ValidoCampos = False

    If Trim(tNombre.Text) = "" Then
        MsgBox "Ingrese el nombre o identificación del valor calculado.", vbExclamation, "Faltan Datos"
        tNombre.SetFocus: Exit Function
    End If
    
    If Trim(tCondicion.Text) = "" Then
        MsgBox "Ingrese el cálculo a realizar para obtener el resultado final.", vbExclamation, "Faltan Datos"
        tCondicion.SetFocus: Exit Function
    End If
        
    ValidoCampos = True
End Function

Private Sub lCondSiNo_Click()
On Error Resume Next

    If cCaminoSiNo.ListIndex <> -1 Then
        prmIdValor = cCaminoSiNo.ItemData(cCaminoSiNo.ListIndex)
        Call Form_Load
    End If
    
End Sub

Private Sub lCondSISI_DblClick()
On Error Resume Next

    If cCaminoSiSi.ListIndex <> -1 Then
        prmIdValor = cCaminoSiSi.ItemData(cCaminoSiSi.ListIndex)
        Call Form_Load
    End If
    
End Sub

Private Sub tCondicion_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errBuscar
    
    If KeyCode = vbKeyF1 Then
        Screen.MousePointer = 11
        
        Cons = "Select VCaCodigo, VCaNombre as 'Valor', VCaDescripcion as 'Descripción' " & _
                   " from ValoresCalculados Where VCaTipo <> " & prmIdValor & _
                   " Order by VCaNombre"
        Dim miLista As New clsListadeAyuda, aItem As Long
        If miLista.ActivarAyuda(cBase, Cons, 4500, 1) > 0 Then aItem = miLista.RetornoDatoSeleccionado(0)
        Me.Refresh
        Set miLista = Nothing
        
        If aItem > 0 Then
            Dim aTexto As String
            Cons = "Select * from ValoresCalculados Where VCaCodigo = " & aItem
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then aTexto = "[" & Trim(RsAux!VCaNombre) & "]" Else aTexto = ""
            RsAux.Close
            
            If aItem <> 0 Then
                With tCondicion
                    aItem = .SelStart
                    .Text = Mid(.Text, 1, .SelStart) & aTexto & Mid(.Text, .SelStart + 1)
                    .SelStart = aItem + Len(aTexto)
                End With
            End If
        End If
        Screen.MousePointer = 0
    End If
    
    Exit Sub
errBuscar:
    clsGeneral.OcurrioError "Error al buscar los valores.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tCondicion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        bEvaluar.SetFocus
    End If
End Sub

Private Sub InicializoControles()

     Cons = "Select VCaCodigo, VCaNombre from ValoresCalculados Where VCaTipo = 2 Order by VCaNombre"
     CargoCombo Cons, cCaminoSiSi
     CargoCombo Cons, cCaminoSiNo
     
     Cons = "Select ConCodigo, ConNombre from CondicionResolucion Order by ConNombre"
     CargoCombo Cons, cResSiSi
     
    bEliminar.Enabled = False
    
End Sub

Private Sub tDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tCondicion.SetFocus
End Sub

Private Sub tNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tDescripcion
End Sub

Private Sub AccionEliminar()

    If MsgBox("Confirma eliminar la condición seleccionada.", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar Condición") = vbNo Then Exit Sub
    
    On Error GoTo errGrabar
    Screen.MousePointer = 11
    
    Cons = "Select * from ValoresCalculados Where VCaCodigo = " & prmIdValor
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then RsAux.Delete
    RsAux.Close
    
    Unload Me
    prmGrabo = True
    Screen.MousePointer = 0
    Exit Sub
    
errGrabar:
    clsGeneral.OcurrioError "Error al eliminar la condición.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tprmSolicitud_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    Select Case KeyCode
        Case vbKeyF2
            EjecutarApp prmPathApp & "Visualizacion de Solicitudes.exe", Trim(tprmSolicitud.Text)
            
        Case vbKeyF1
                 
                 Cons = "SELECT TOP 20 SolCodigo, CliTipo, CliCIRuc, RTrim(SolComentarioR), Sum(RSoValorCuota) as TotCuota, IsNull(Sum(RSoValorEntrega), 0) as TotEnt, Sum(TCuCantidad * RSoValorCuota) as TotFin, Count(RSoSolicitud) as Artículos " & _
                            " From Solicitud, Cliente, RenglonSolicitud, TipoCuota " & _
                            " Where SolCliente = CliCodigo And SolCodigo = RSoSolicitud And TCuCodigo = RSoTipoCuota " & _
                            " GROUP BY SolCodigo, CliTipo, CliCIRuc, SolComentarioR " & _
                            " ORDER BY SolCodigo DESC"
                
                Dim miLista As New clsListadeAyuda
                If miLista.ActivarAyuda(cBase, Cons, 7500, 0) > 0 Then tprmSolicitud.Text = miLista.RetornoDatoSeleccionado(0)
                Me.Refresh
                Set miLista = Nothing

    End Select
    
End Sub

Private Sub txtTextoSiSi_GotFocus()
    With txtTextoSiSi
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTextoSiSi_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then cCaminoSiSi.SetFocus
End Sub
