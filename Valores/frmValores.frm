VERSION 5.00
Begin VB.Form frmValores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valores Calculados"
   ClientHeight    =   5205
   ClientLeft      =   2490
   ClientTop       =   4425
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
   Icon            =   "frmValores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7575
   Begin VB.CheckBox chkGrabarHistoria 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "¿Grabar historia?"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton bGrabarCerrar 
      Caption         =   "Grabar/Salir"
      Height          =   315
      Left            =   5160
      TabIndex        =   21
      Top             =   4860
      Width           =   1095
   End
   Begin VB.TextBox tResultado 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   4440
      Width           =   6495
   End
   Begin VB.CommandButton bEliminar 
      Caption         =   "Eliminar"
      Height          =   315
      Left            =   2760
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4860
      Width           =   1095
   End
   Begin VB.CommandButton bCancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   6360
      TabIndex        =   12
      Top             =   4860
      Width           =   1095
   End
   Begin VB.CommandButton bGrabar 
      Caption         =   "Grabar"
      Height          =   315
      Left            =   3960
      TabIndex        =   11
      Top             =   4860
      Width           =   1095
   End
   Begin VB.TextBox tprmSolicitud 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   14
      Text            =   "3200"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton bPrevia 
      Caption         =   "&Vista Previa"
      Height          =   315
      Left            =   6360
      TabIndex        =   10
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton bEvaluar 
      Caption         =   "&Evaluar"
      Height          =   315
      Left            =   5160
      TabIndex        =   9
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox tCondicion 
      Appearance      =   0  'Flat
      Height          =   2265
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1740
      Width           =   7335
   End
   Begin VB.TextBox tFormato 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3720
      MaxLength       =   15
      TabIndex        =   8
      Top             =   4080
      Width           =   1335
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
   Begin VB.Label lCabezal2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      Height          =   45
      Left            =   60
      TabIndex        =   18
      Top             =   4760
      Width           =   4875
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingreso del Valor Calculado"
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
      TabIndex        =   17
      Top             =   120
      Width           =   2835
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "frmValores.frx":030A
      Top             =   20
      Width           =   480
   End
   Begin VB.Label lCabezal 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      Height          =   45
      Left            =   0
      TabIndex        =   16
      Top             =   540
      Width           =   4875
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "[prmSolicitud]:"
      Height          =   255
      Left            =   5460
      TabIndex        =   15
      Top             =   165
      Width           =   1035
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Resultado:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4470
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
      Caption         =   "&Formato:"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   4115
      Width           =   735
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
Attribute VB_Name = "frmValores"
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

Private Sub bGrabar_Click()
    AccionGrabar
End Sub

Private Sub bGrabarCerrar_Click()
    AccionGrabar cExit:=True
End Sub

Private Sub bPrevia_Click()
    AccionVistaPrevia
End Sub


Private Sub chkGrabarHistoria_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then tFormato.SetFocus
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    chkGrabarHistoria.BackColor = Me.BackColor
    
    'Center form
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    prmGrabo = False
    bEliminar.Enabled = False
    If prmIdValor <> 0 Then
        CargoDatosValor
        bEliminar.Enabled = True
    End If
    
    Screen.MousePointer = 0
    
End Sub

Private Sub AccionEvaluar()
    On Error GoTo errEvaluar
    
    tResultado.Text = ""
    Screen.MousePointer = 11
    
    Dim mOBJ As New clsResolAuto
    Dim aResult As Variant, aExp As String
    
    aResult = mOBJ.EvaluoCondicion(cBase, Val(tprmSolicitud.Text), Trim(tCondicion.Text), aExp)
    Set mOBJ = Nothing
    
    If Trim(tFormato.Text) <> "" Then aResult = Format(aResult, Trim(tFormato.Text))
    tResultado.Text = aResult
    
    Screen.MousePointer = 0
    
    Exit Sub
errEvaluar:
    clsGeneral.OcurrioError "Error al evaluar la condición", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionVistaPrevia()
    On Error GoTo errEvaluar
    
    If UCase(Mid(tCondicion, 1, 6)) <> "SELECT" Then
        MsgBox "Sólo se pueden ver en vista previa a las consultas SQL.", vbInformation, "Falta consulta SQL"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    Dim mOBJ As New clsResolAuto
    mOBJ.EvaluoCondicion cBase, Val(tprmSolicitud.Text), Trim(tCondicion.Text), Cons
    Set mOBJ = Nothing
    
    Dim miSql As New clsListadeAyuda
    miSql.ActivoListaAyudaSQL cBase, Cons
    Set miSql = Nothing
    
    Screen.MousePointer = 0
    Exit Sub
errEvaluar:
    clsGeneral.OcurrioError "Error al evaluar la condición", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionGrabar(Optional cExit As Boolean = False)

    If Not ValidoCampos Then Exit Sub
    If MsgBox("Confirma grabar los datos para el valor calculado: " & Trim(tNombre.Text), vbQuestion + vbYesNo, "Grabar") = vbNo Then Exit Sub
    
    On Error GoTo errGrabar
    Screen.MousePointer = 11
    
    Cons = "Select * from ValoresCalculados Where VCaCodigo = " & prmIdValor
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then RsAux.Edit Else RsAux.AddNew
    
    RsAux!VCaNombre = Trim(tNombre.Text)
    RsAux!VCaTipo = 1
    If Trim(tDescripcion.Text) <> "" Then RsAux!VCaDescripcion = Trim(tDescripcion.Text) Else RsAux!VCaDescripcion = Null
    RsAux!VCaTexto = Trim(tCondicion.Text)
    If Trim(tFormato.Text) <> "" Then RsAux!VCaFormato = Trim(tFormato.Text) Else RsAux!VCaFormato = Null
    
    RsAux("VCaGrabarValor") = (IIf(chkGrabarHistoria.Value, 1, 0))
    
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
    clsGeneral.OcurrioError "Error al grabar el valor calculado.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoDatosValor()

    On Error GoTo errCargar
    Screen.MousePointer = 11
    
    Cons = "Select * from ValoresCalculados Where VCaCodigo = " & prmIdValor
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
    
        tNombre.Text = Trim(RsAux!VCaNombre)
        If Not IsNull(RsAux!VCaDescripcion) Then tDescripcion.Text = Trim(RsAux!VCaDescripcion) Else tDescripcion.Text = ""
        tCondicion.Text = Trim(RsAux!VCaTexto)
        If Not IsNull(RsAux!VCaFormato) Then tFormato.Text = Trim(RsAux!VCaFormato) Else tFormato.Text = ""
        If Not IsNull(RsAux("VCaGrabarValor")) Then
            chkGrabarHistoria.Value = IIf(RsAux("VCaGrabarValor") = 0, 0, 1)
        End If
    
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

Private Sub tCondicion_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errBuscar
    
    If KeyCode = vbKeyF1 Then
        Screen.MousePointer = 11
        
        Cons = "Select VCaCodigo, VCaNombre as 'Valor', VCaDescripcion as 'Descripción' " & _
                    " from ValoresCalculados Where VCaTipo <> " & prmIdValor & _
                    " Order by VCaNombre"
        Dim miLista As New clsListadeAyuda, aItem As Long
        If miLista.ActivarAyuda(cBase, Cons, 4500, 1) > 0 Then
            aItem = miLista.RetornoDatoSeleccionado(0)
        End If
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
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        chkGrabarHistoria.SetFocus
    End If
End Sub

Private Sub tDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tCondicion.SetFocus
End Sub

Private Sub tFormato_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bGrabar.SetFocus
End Sub

Private Sub tNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tDescripcion
End Sub

Private Sub AccionEliminar()

    If MsgBox("Confirma eliminar el valor calculado.", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar Valor") = vbNo Then Exit Sub
    
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
    clsGeneral.OcurrioError "Error al eliminar el valor calculado.", Err.Description
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

