VERSION 5.00
Begin VB.Form frmComentario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comentario de inventario"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmComentario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bSalir 
      Caption         =   "&Salir"
      Height          =   315
      Left            =   5460
      TabIndex        =   7
      Top             =   2220
      Width           =   975
   End
   Begin VB.CommandButton bGrabar 
      Caption         =   "&Grabar"
      Height          =   315
      Left            =   4380
      TabIndex        =   6
      Top             =   2220
      Width           =   975
   End
   Begin VB.TextBox tUsuario 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1020
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2220
      Width           =   1155
   End
   Begin VB.TextBox tComentario 
      Appearance      =   0  'Flat
      Height          =   1635
      Left            =   1020
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frmComentario.frx":030A
      Top             =   540
      Width           =   5415
   End
   Begin VB.TextBox tArticulo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1020
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   180
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2220
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Comentario:"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   540
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Artículo:"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   915
   End
End
Attribute VB_Name = "frmComentario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_Art As Long
Dim sTerminal As String

Public Property Let prmArticulo(ByVal lID As Long)
    m_Art = lID
End Property
Private Sub bGrabar_Click()
    AccionGrabar
End Sub

Private Sub bSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
    ObtengoSeteoForm Me, 500, 500
    cons = "Select * From Terminal Where TerCodigo = " & paCodigoDeTerminal
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If rsAux.EOF Then
        MsgBox "No se obtuvo la terminal.", vbExclamation, "ATENCIÓN"
        sTerminal = "S/Term"
    Else
        sTerminal = Trim(rsAux!TerNombre)
    End If
    rsAux.Close
    tArticulo.Text = "": tComentario.Text = "": tUsuario.Text = ""
    If m_Art > 0 Then BuscoArticuloPorCodigo 0, m_Art
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    GuardoSeteoForm Me
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = "0"
    LimpioDatos
End Sub

Private Sub tArticulo_GotFocus()
    With tArticulo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub tArticulo_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrAP
    
    If KeyAscii = vbKeyReturn Then
        If Val(tArticulo.Tag) > 0 And tArticulo.Text <> "" Then tComentario.SetFocus: Exit Sub
        Screen.MousePointer = 11
        If Trim(tArticulo.Text) <> "" Then
            If IsNumeric(tArticulo.Text) Then
                BuscoArticuloPorCodigo tArticulo.Text
            Else
                BuscoArticuloPorNombre tArticulo.Text
            End If
            If Val(tArticulo.Tag) > 0 Then tComentario.SetFocus
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrAP:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub BuscoArticuloPorCodigo(ByVal CodArticulo As Long, Optional ID As Long = 0)
'Atención el mapeo de error lo hago antes de entrar al procedimiento
    
    Screen.MousePointer = 11
    If ID = 0 Then
        cons = "Select * From Articulo Where ArtCodigo = " & CodArticulo
    Else
        cons = "Select * From Articulo Where ArtID = " & ID
    End If
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If rsAux.EOF Then
        rsAux.Close
        tArticulo.Tag = "0"
        MsgBox "No existe un artículo que posea ese código.", vbExclamation, "ATENCIÓN"
    Else
        tArticulo.Text = Format(rsAux!ArtCodigo, "#,000,000") & " " & Trim(rsAux!ArtNombre)
        tArticulo.Tag = rsAux!ArtID
        rsAux.Close
    End If
    Screen.MousePointer = 0

End Sub

Private Sub BuscoArticuloPorNombre(NomArticulo As String)
'Atención el mapeo de error lo hago antes de entrar al procedimiento
Dim Resultado As Long
Dim objAyuda As clsListadeAyuda
    
    Screen.MousePointer = 11
    cons = "Select Count(*) From Articulo " _
        & " Where ArtNombre LIKE '" & Replace(NomArticulo, " ", "%") & "%'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not IsNull(rsAux(0)) Then
        Resultado = rsAux(0)
    Else
        Resultado = 0
    End If
    rsAux.Close
    
    If Resultado = 0 Then
        MsgBox "No hay datos para el filtro ingresado.", vbInformation, "ATENCIÓN"
    Else
        If Resultado = 1 Then
            cons = "Select * From Articulo " _
                & " Where ArtNombre LIKE '" & Replace(NomArticulo, " ", "%") & "%'"
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurReadOnly)
            Resultado = rsAux!ArtCodigo
            rsAux.Close
        Else
            cons = "Select Código = ArtCodigo, Nombre = ArtNombre from Articulo" _
                & " Where ArtNombre LIKE '" & Replace(NomArticulo, " ", "%") & "%'" _
                & " Order By ArtNombre"
            
            Set objAyuda = New clsListadeAyuda
            If objAyuda.ActivarAyuda(cBase, cons, 4500, 0, "Ayuda de Artículos") > 0 Then
                Screen.MousePointer = 11
                Resultado = objAyuda.RetornoDatoSeleccionado(0)
            Else
                Resultado = 0
            End If
            Set objAyuda = Nothing
        End If
    End If
    If Resultado > 0 Then BuscoArticuloPorCodigo Resultado
    Screen.MousePointer = 0
    
End Sub

Private Sub LimpioDatos()
    tComentario.Text = ""
End Sub

Private Sub GraboEnBD()
Dim sAntes As String, sEncabezado As String

    If Trim(tComentario.Text) = "" Then MsgBox "Ingrese el comentario.", vbExclamation, "ATENCIÓN": Exit Sub
    If Val(tUsuario.Tag) = 0 Then MsgBox "Ingrese su dígito de usuario.", vbExclamation, "ATENCIÓN": tUsuario.SetFocus: Exit Sub
    Screen.MousePointer = 11
    On Error GoTo errBT
    '-----------------------------------------------------
    sEncabezado = String(70, "-") & vbCrLf & "Usuario: " & Trim(tUsuario.Text) & " Fecha: " & Format(Now, "dd/mm/yy hh:mm") _
                                        & " Terminal:" & sTerminal
    cBase.BeginTrans
    On Error GoTo errRB
    cons = "Select * From ComentarioInventario Where CInArticulo = " & Val(tArticulo.Tag)
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If rsAux.EOF Then
        rsAux.AddNew
        rsAux!CInTexto = sEncabezado & Trim(tComentario.Text)
    Else
        rsAux.Edit
        sAntes = rsAux!CInTexto
        rsAux!CInTexto = sAntes & vbCrLf & sEncabezado & vbCrLf & Trim(tComentario.Text)
    End If
    rsAux!CInArticulo = Val(tArticulo.Tag)
    rsAux.Update
    rsAux.Close
    cBase.CommitTrans
    '-----------------------------------------------------
    tComentario.Text = "": tArticulo.Text = ""
    tArticulo.SetFocus
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
errRB:
    Resume RollB

End Sub

Private Sub tUsuario_Change()
    tUsuario.Tag = ""
End Sub

Private Sub tUsuario_GotFocus()
    With tUsuario
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Val(tUsuario.Tag) > 0 Then
            If Trim(tComentario.Text) <> "" Then AccionGrabar
        Else
            If Trim(tUsuario.Text) <> "" Then
                If IsNumeric(tUsuario.Text) Then
                    cons = "Select * From Usuario Where UsuDigito = " & Val(tUsuario.Text)
                    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                    If rsAux.EOF Then
                        rsAux.Close
                        MsgBox "Dígito incorrecto.", vbExclamation, "ATENCIÓN"
                    Else
                        tUsuario.Text = Trim(rsAux!UsuIdentificacion)
                        tUsuario.Tag = rsAux!UsuCodigo
                        rsAux.Close
                        If Trim(tComentario.Text) <> "" Then AccionGrabar
                    End If
                Else
                    MsgBox "Debe ingresar su dígito.", vbExclamation, "ATENCIÓN"
                End If
            End If
        End If
    End If
    
End Sub

Private Sub AccionGrabar()
    If Val(tArticulo.Tag) = 0 Then MsgBox "Ingrese un artículo.", vbExclamation, "ATENCIÓN": Exit Sub
    If MsgBox("¿Confirma almacenar el comentario?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then
        GraboEnBD
    End If
End Sub
