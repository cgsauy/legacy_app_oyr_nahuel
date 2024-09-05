VERSION 5.00
Begin VB.Form frmFiltroEtiqueta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregar Etiquetas según filtro"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3990
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton bCancel 
      Caption         =   "C&ancelar"
      Height          =   375
      Left            =   2880
      TabIndex        =   14
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton bAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox tCantEtiqueta 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Text            =   "1"
      Top             =   1680
      Width           =   495
   End
   Begin VB.VScrollBar vscCantidad 
      Height          =   285
      Left            =   1920
      Min             =   1
      TabIndex        =   10
      Top             =   1680
      Value           =   1
      Width           =   255
   End
   Begin VB.ComboBox cQueEtiqueta 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox tLista 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox tMarca 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox tTipo 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox tFecha 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "&Cantidad:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "¿&Qué etiqueta?:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Lista:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Marca:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Tipo de Artículo:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Precio modificado después del:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmFiltroEtiqueta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private colArticulo As New Collection
Private m_CantE As Long, m_QueEtiqueta As Integer

Public Property Get prmHayDatos() As Boolean
    If colArticulo Is Nothing Then
        prmHayDatos = False
    Else
        If colArticulo.Count = 0 Then
            prmHayDatos = False
        Else
            prmHayDatos = True
        End If
    End If
End Property
Public Property Get prmCantResultado() As Long
    prmCantResultado = colArticulo.Count
End Property

Public Property Get prmIDResultado(ByVal lPos As Long) As Long
    If colArticulo Is Nothing Then Exit Sub
    If lPos > colArticulo.Count Or lPos < 1 Then Exit Sub
    prmIDResultado = colArticulo(lPos)
End Property

Public Property Get prmCantidad() As Long
    prmCantidad = m_CantE
End Property

Public Property Get prmQueEtiqueta() As Integer
    prmQueEtiqueta = m_QueEtiqueta
End Property

Private Sub bAplicar_Click()
'Consulto dados los filtros
    m_QueEtiqueta = -1
    m_CantE = 0
    Set colArticulo = Nothing
    If tFecha.Text <> "" Or Val(tTipo.Tag) > 0 Or Val(tMarca.Tag) > 0 Or Val(tLista.Tag) > 0 Then
    
        Cons = "Select Distinct(ArtCodigo) From " & _
                        "HistoriaPrecio Precios, " & _
                        "Articulo, ArticuloFacturacion, TipoCuota " & _
                " Where Precios.HPrArticulo = ArtID"

        If IsDate(tFecha.Text) Then
            Cons = Cons & " And HPrVigencia >= '" & Format(tFecha.Text, "mm/dd/yyyy 00:00:00") & "'"
        End If
        If Val(tTipo.Tag) > 0 Then Cons = Cons & " And ArtTipo = " & Val(tTipo.Tag)
        If Val(tMarca.Tag) > 0 Then Cons = Cons & " And ArtMarca = " & Val(tMarca.Tag)
        If Val(tLista.Tag) > 0 Then Cons = Cons & " And AFaLista = " & Val(tLista.Tag)
        

        Cons = Cons & " And ArtId = AFaArticulo " & _
                " And ArtEnUso = 1" & _
                " And Precios.HPrMoneda = " & paMonedaPesos & _
                " And Precios.HPrHabilitado = 1" & _
                " And Precios.HPrTipoCuota = TipoCuota.TCuCodigo"
        Cons = Cons & _
                " And TipoCuota.TCuVencimientoE Is Null " & _
                " And TCuEspecial = 0 And TCuDeshabilitado Is Null"
        
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If RsAux.EOF Then
            MsgBox "No hay datos para los filtros ingresados.", vbInformation, "ATENCIÓN"
            RsAux.Close
        Else
            Do While Not RsAux.EOF
                colArticulo.Add CStr(RsAux(0))
                RsAux.MoveNext
            Loop
            RsAux.Close
            m_QueEtiqueta = cQueEtiqueta.ListIndex
            m_CantE = CLng(tCantEtiqueta.Text)
            Unload Me
        End If
    Else
        MsgBox "No se ingresaron filtros.", vbExclamation, "ATENCIÓN"
    End If
    
End Sub

Private Sub bCancel_Click()
    Unload Me
End Sub

Private Sub cQueEtiqueta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bAplicar.SetFocus
End Sub

Private Sub Form_Load()
    With cQueEtiqueta
        .Clear
        .AddItem "Ambas"
        .AddItem "Normal (chica)"
        .AddItem "Según tabla"
        .AddItem "Vidriera (grande)"
        .ListIndex = 2
    End With
End Sub

Private Sub tCantEtiqueta_GotFocus()
    With tCantEtiqueta
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    If Val(tCantEtiqueta.Text) = 0 Then tCantEtiqueta.Text = vscCantidad.Value
End Sub

Private Sub tCantEtiqueta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tCantEtiqueta.Text) Then
            If Val(tCantEtiqueta.Text) < 1 Then tCantEtiqueta.Text = vscCantidad.Value
            vscCantidad.Value = Val(tCantEtiqueta.Text)
            cQueEtiqueta.SetFocus
        Else
            MsgBox "Formato incorrecto.", vbExclamation, "ATENCIÓN"
            tCantEtiqueta.Text = vscCantidad.Value
        End If
    End If
End Sub

Private Sub tCantEtiqueta_LostFocus()
    If Not IsNumeric(tCantEtiqueta.Text) Then
        tCantEtiqueta.Text = vscCantidad.Value
    Else
        If Val(tCantEtiqueta.Text) < 1 Then tCantEtiqueta.Text = vscCantidad.Value
    End If
End Sub

Private Sub tFecha_GotFocus()
    With tFecha
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If IsDate(tFecha.Text) Then
            tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy")
            tTipo.SetFocus
        Else
            If tFecha.Text <> "" Then
                MsgBox "Formato de fecha incorrecto.", vbExclamation, "ATENCIÓN"
                tFecha.Text = ""
            Else
                tTipo.SetFocus
            End If
        End If
    End If
End Sub

Private Sub tFecha_LostFocus()
On Error Resume Next
    If IsDate(tFecha.Text) Then
        tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy")
    Else
        tFecha.Text = ""
    End If
End Sub

Private Sub tLista_Change()
    If Val(tLista.Tag) > 0 Then tLista.Tag = ""
End Sub

Private Sub tLista_KeyPress(KeyAscii As Integer)
On Error Resume Next
    
    If KeyAscii = vbKeyReturn Then
        If Trim(tLista.Text) = "" Or Val(tLista.Tag) > 0 Then
            tCantEtiqueta.SetFocus
        Else
            If Trim(tLista.Text) <> "" Then
                Cons = "Select LDPCodigo, LDPDescripcion as 'Lista de Precios' From ListasDePrecios " _
                    & " Where  LDPDescripcion Like '" & Replace(tLista.Text, " ", "%") & "%'" _
                    & "Order by LDPDescripcion"
                ListaAyuda Cons, tLista, "Listas de Precios"
                If Val(tLista.Tag) > 0 Then tCantEtiqueta.SetFocus
            End If
        End If
    End If

End Sub

Private Sub tMarca_Change()
    If Val(tMarca.Text) > 0 Then tMarca.Tag = ""
End Sub

Private Sub tMarca_GotFocus()
    With tMarca
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tMarca_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = vbKeyReturn Then
        If Trim(tMarca.Text) = "" Or Val(tMarca.Tag) > 0 Then
            tLista.SetFocus
        Else
            If Trim(tMarca.Text) <> "" Then
                Cons = "Select MarCodigo, MarNombre as 'Marca' From Marca " _
                    & " Where MarNombre Like '" & Replace(tMarca.Text, " ", "%") & "%'" _
                    & "Order by MarNombre"
                ListaAyuda Cons, tMarca, "Lista de Marcas"
                If Val(tMarca.Tag) > 0 Then tLista.SetFocus
            End If
        End If
    End If

End Sub

Private Sub tTipo_Change()
    If Val(tTipo.Tag) > 0 Then tTipo.Tag = ""
End Sub

Private Sub tTipo_GotFocus()
    With tTipo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tTipo_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = vbKeyReturn Then
        If Trim(tTipo.Text) = "" Or Val(tTipo.Tag) > 0 Then
            tMarca.SetFocus
        Else
            If Trim(tTipo.Text) <> "" Then
                Cons = "Select TipCodigo, TipNombre as 'Tipo' From Tipo " _
                    & " Where TipNombre Like '" & Replace(tTipo.Text, " ", "%") & "%'" _
                    & "Order by TipNombre"
                ListaAyuda Cons, tTipo, "Tipos de Artículos"
                If Val(tTipo.Tag) > 0 Then tMarca.SetFocus
            End If
        End If
    End If
    
End Sub

Private Sub ListaAyuda(ByVal sCons As String, tControl As Control, ByVal sTitulo As String)
Dim objLista As New clsListadeAyuda
    '1ero hago cons. para ver cantidad.
    
    Set RsAux = cBase.OpenResultset(sCons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        MsgBox "No se encontraron datos para el filtro ingresado.", vbExclamation, "ATENCIÓN"
        RsAux.Close
    Else
        RsAux.MoveNext
        If RsAux.EOF Then
            RsAux.MoveFirst
            tControl.Text = Trim(RsAux(1))
            tControl.Tag = RsAux(0)
            RsAux.Close
        Else
            RsAux.Close
            If objLista.ActivarAyuda(cBase, sCons, 5000, 1, sTitulo) > 0 Then
                tControl.Text = objLista.RetornoDatoSeleccionado(1)
                tControl.Tag = objLista.RetornoDatoSeleccionado(0)
            End If
        End If
    End If
    Set objLista = Nothing
                
End Sub

Private Sub vscCantidad_Change()
    tCantEtiqueta.Text = vscCantidad.Value
End Sub
