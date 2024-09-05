VERSION 5.00
Begin VB.Form ComVtaReal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comentario de Ventas Reales"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ComVtaReal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fBotones 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   4815
      Begin VB.CommandButton bGrabar 
         Height          =   310
         Left            =   3600
         Picture         =   "ComVtaReal.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Grabar Comentario. [Ctrl+G]"
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bEliminar 
         Height          =   310
         Left            =   3960
         Picture         =   "ComVtaReal.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Eliminar comentario. [Ctrl+E]"
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   4320
         Picture         =   "ComVtaReal.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Salir. [Ctrl+X]"
         Top             =   0
         Width           =   310
      End
   End
   Begin VB.TextBox tComentario 
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   1200
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox tMes 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1200
      MaxLength       =   12
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox tArticulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Comentario:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Mes:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Artículo:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "ComVtaReal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private iSeleccionado As Long
Private strMes As String
Public Property Get pSeleccionado() As Long
    pSeleccionado = iSeleccionado
End Property
Public Property Let pSeleccionado(Codigo As Long)
    iSeleccionado = Codigo
End Property
Public Property Get pMes() As String
    pMes = strMes
End Property
Public Property Let pMes(Mes As String)
    strMes = Mes
End Property
Private Sub bCancelar_Click()
    Unload Me
End Sub
Private Sub bEliminar_Click()
On Error GoTo ErrBE
    If tArticulo.Tag = "" And Not IsDate(tMes.Text) Then MsgBox "No hay información a eliminar.", vbExclamation, "ATENCIÓN": Exit Sub
    RelojA
    Cons = "Select * From CodigoTexto" _
        & " Where Clase = " & tArticulo.Tag _
        & " And Texto2 = '" & Format("01/" & tMes.Text, sqlFormatoF) & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsAux.EOF Then
        RsAux.Delete
    End If
    RsAux.Close
    RelojD
    Exit Sub
ErrBE:
    MensajeError "Ocurrió un error al buscar el comentario.", Err.Description
    RelojD
End Sub
Private Sub bGrabar_Click()
    If tArticulo.Tag = "" And Not IsDate(tMes.Text) Then MsgBox "No hay información a eliminar.", vbExclamation, "ATENCIÓN": Exit Sub
    If MsgBox("¿Desea almacenar el comentario ingresado?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then GraboComentario
End Sub
Private Sub Form_Activate()
    RelojD
    If tArticulo.Text <> "" Then tMes.SetFocus Else tArticulo.SetFocus
    Me.Refresh
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyX
                Unload Me
            Case vbKeyG
                bGrabar_Click
            Case vbKeyE
                bEliminar_Click
        End Select
    Else
        Select Case KeyCode
            Case vbKeyEscape
                Unload Me
        End Select
    End If
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad
    If iSeleccionado > 0 Then
        BuscoArticuloPorID iSeleccionado
        If Trim(strMes) <> "" Then tMes.Text = strMes: BuscoComentario
    End If
    Exit Sub
ErrLoad:
    MensajeError "Ocurrió un error al inicializar el formulario.", Err.Description
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If VtasEstimadas.WindowState <> vbMinimized Then VtasEstimadas.SetFocus
End Sub
Private Sub Label1_Click()
    Foco tArticulo
End Sub
Private Sub Label2_Click()
    Foco tMes
End Sub
Private Sub Label3_Click()
    Foco tComentario
End Sub
Private Sub tArticulo_Change()
    tArticulo.Tag = ""
    tMes.Text = ""
    tComentario.Text = ""
End Sub
Private Sub tArticulo_GotFocus()
    With tArticulo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub tArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(tArticulo.Text) <> "" Then
        tMes.Text = "": tComentario.Text = ""
        If IsNumeric(tArticulo.Text) Then
            BuscoArticuloPorCodigo CLng(tArticulo.Text)
        Else
            BuscoArticuloPorNombre
        End If
        If Trim(tArticulo.Tag) <> "" Then tMes.SetFocus
    End If
End Sub
Private Sub BuscoArticuloPorCodigo(Articulo As Long)
On Error GoTo ErrBAPC
    RelojA
    Cons = "Select ArtID, ArtNombre, ArtHabilitado From Articulo Where ArtCodigo = " & CLng(Articulo)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If RsAux.EOF Then
        RelojD
        MsgBox "No existe un artículo con ese código, o el mismo fue eliminado.", vbInformation, "ATENCIÓN"
        tComentario.Text = "": tMes.Text = ""
    Else
        If Not IsNull(RsAux!ArtHabilitado) Then
            If Trim(UCase(RsAux!ArtHabilitado)) = "S" Then
                tArticulo.Text = Trim(RsAux!ArtNombre): tArticulo.Tag = RsAux!ArtID
            Else
                MsgBox "El artículo seleccionado no esta habilitado.", vbInformation, "ATENCIÓN"
            End If
        Else
            MsgBox "El artículo seleccionado no esta habilitado.", vbInformation, "ATENCIÓN"
        End If
    End If
    RsAux.Close
    RelojD
    Exit Sub
ErrBAPC:
    MensajeError "Ocurrió un error al buscar el artículo por código.", Err.Description
    RelojD
End Sub
Private Sub BuscoArticuloPorID(Articulo As Long)
On Error GoTo ErrBAPI
    RelojA
    Cons = "Select ArtID, ArtNombre, ArtHabilitado From Articulo Where ArtID = " & CLng(Articulo)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!ArtHabilitado) Then
            If Trim(UCase(RsAux!ArtHabilitado)) = "S" Then
                tArticulo.Text = Trim(RsAux!ArtNombre): tArticulo.Tag = RsAux!ArtID
            Else
                MsgBox "El artículo seleccionado no esta habilitado.", vbInformation, "ATENCIÓN"
            End If
        Else
            MsgBox "El artículo seleccionado no esta habilitado.", vbInformation, "ATENCIÓN"
        End If
    End If
    RsAux.Close
    RelojD
    Exit Sub
ErrBAPI:
    MensajeError "Ocurrió un error al buscar el artículo por código.", Err.Description
    RelojD
End Sub
Private Sub BuscoArticuloPorNombre()
    Cons = "Select ArtCodigo, 'Código' = ArtCodigo, Nombre = ArtNombre" _
        & " From Articulo" _
        & " Where ArtNombre LIKE '" & Trim(tArticulo.Text) & "%'"
    PresentoListaDeAyuda Cons
End Sub
Private Sub PresentoListaDeAyuda(strConsulta As String)
On Error GoTo ErrPLDA
Dim Resultado As String
    
    RelojA
    'Limpio los valores del textbox.
    tArticulo.Tag = "": tArticulo.Text = ""
    
    Dim sqlAyuda As New clsListadeAyuda
    sqlAyuda.ActivarAyuda cBase, Cons, 4000, 0, "Ayuda"
    'Obtengo si hay seleccionado.---------------
    Resultado = sqlAyuda.RetornoDatoSeleccionado(0)
    'Destruyo la clase.------------------------------
    Set sqlAyuda = Nothing
    RelojA
    If Resultado <> "" Then
        If IsNumeric(Resultado) Then
           BuscoArticuloPorCodigo CLng(Resultado)
        Else
            RelojD
            MsgBox "Se espera que se retorne el código de artículo.", vbInformation, "ATENCIÓN"
        End If
    End If
    RelojD
    Exit Sub
ErrPLDA:
    RelojD
    MensajeError "Ocurrió un error al presentar la lista de ayuda.", Err.Description
End Sub
Private Sub tComentario_GotFocus()
    With tComentario
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bGrabar.SetFocus
End Sub
Private Sub tMes_GotFocus()
    With tMes
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub tMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsDate(tMes.Text) And Trim(tArticulo.Tag) <> "" Then BuscoComentario: tComentario.SetFocus
    End If
End Sub
Private Sub BuscoComentario()
On Error GoTo ErrBC
    RelojA
    Cons = "Select * From ComentarioVentaReal" _
        & " Where CVRArticulo = " & tArticulo.Tag _
        & " And CVRFecha = '" & Format("01/" & tMes.Text, sqlFormatoF) & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsAux.EOF Then
        tComentario.Text = Trim(RsAux!CVRComentario)
    Else
        tComentario.Text = ""
    End If
    RsAux.Close
    RelojD
    Exit Sub
ErrBC:
    MensajeError "Ocurrió un error al buscar el comentario.", Err.Description
    RelojD
End Sub
Private Sub GraboComentario()
On Error GoTo ErrBC
    RelojA
    Cons = "Select * From CodigoTexto" _
        & " Where Tipo = " & TipoEnCodTexto _
        & " And Clase = " & tArticulo.Tag _
        & " And Texto2 = '" & Format("01/" & tMes.Text, sqlFormatoF) & "'"
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        If Trim(SacoEnter(tComentario.Text)) <> "" Then
            RsAux.AddNew
            RsAux!Tipo = TipoEnCodTexto
            RsAux!Clase = tArticulo.Tag
            RsAux!Texto2 = Format("01/" & tMes.Text, sqlFormatoF)
            RsAux!Texto = Trim(SacoEnter(tComentario.Text))
            RsAux.Update
        End If
    Else
        If Trim(SacoEnter(tComentario.Text)) <> "" Then
            RsAux.Edit
            RsAux!Texto = Trim(SacoEnter(tComentario.Text))
            RsAux.Update
        Else
            RsAux.Delete
        End If
    End If
    RsAux.Close
    RelojD
    Exit Sub
ErrBC:
    MensajeError "Ocurrió un error al buscar el comentario.", Err.Description
    RelojD
End Sub
Private Sub tMes_LostFocus()
    If IsDate(tMes.Text) Then tMes.Text = Format(tMes.Text, "Mmm/yyyy")
End Sub
