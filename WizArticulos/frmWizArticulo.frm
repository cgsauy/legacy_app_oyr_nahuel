VERSION 5.00
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form frmWizArticulo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asistente de Artículos"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   ControlBox      =   0   'False
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
   ScaleHeight     =   5040
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picWizard 
      Height          =   2535
      Index           =   1
      Left            =   5640
      ScaleHeight     =   2475
      ScaleWidth      =   5955
      TabIndex        =   19
      Top             =   4440
      Width           =   6015
   End
   Begin VB.PictureBox picWizard 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3615
      Index           =   0
      Left            =   0
      ScaleHeight     =   3615
      ScaleWidth      =   6375
      TabIndex        =   18
      Top             =   720
      Width           =   6375
      Begin VB.TextBox tCodigo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox tOrigen 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   15
         Top             =   2880
         Width           =   2655
      End
      Begin VB.TextBox tMarca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox tTipo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox tProveedor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   13
         Top             =   2520
         Width           =   5055
      End
      Begin VB.TextBox tNombre 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   11
         Top             =   2160
         Width           =   5055
      End
      Begin VB.TextBox tDescripcion 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox tParecidoA 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   5055
      End
      Begin AACombo99.AACombo cIVA 
         Height          =   315
         Left            =   1200
         TabIndex        =   21
         Top             =   3240
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "&Codigo:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "&Origen:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "&I.V.A.:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Pro&veedor:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nombre:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Descripción:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Marca:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lParecido 
         BackStyle       =   0  'Transparent
         Caption         =   "&Parecido a:"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   6735
      TabIndex        =   16
      Top             =   0
      Width           =   6735
      Begin VB.Label lLogo 
         BackStyle       =   0  'Transparent
         Caption         =   "Selección"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   17
         Top             =   120
         Width           =   5415
      End
      Begin VB.Image imglogo 
         Height          =   480
         Index           =   0
         Left            =   120
         Picture         =   "frmWizArticulo.frx":0000
         Top             =   120
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmWizArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public prmID As Long

Private sAbrevTipo As String
Private Sub db_ShowHelpList(ByVal sQuery As String, ByVal textB As TextBox, ByVal sTitulo As String, Optional iColHide As Integer = 1)
On Error GoTo errSH
    
    Dim objHelp As New clsListadeAyuda
    With objHelp
        If .ActivarAyuda(cBase, sQuery, , iColHide, sTitulo) > 0 Then
            
            If textB.Name = "tTipo" Then
                textB.Text = .RetornoDatoSeleccionado(2)
                sAbrevTipo = .RetornoDatoSeleccionado(1)
            Else
                textB.Text = .RetornoDatoSeleccionado(1)
            End If
            textB.Tag = .RetornoDatoSeleccionado(0)
        End If
    End With
    Set objHelp = Nothing
    Exit Sub
errSH:
    clsGeneral.OcurrioError "Error al activar la lista de ayuda.", Err.Description
End Sub

Private Sub db_FindTipoMarca(ByVal sFind As String, ByVal textB As TextBox)
On Error GoTo errFT
Dim sAux As String
    
    Screen.MousePointer = 11
    If textB.Name = "tTipo" Then
        sAux = " Tipo"
        If IsNumeric(sFind) Then
            Cons = "Select  TipCodigo, TipAbreviacion, TipNombre as 'Nombre' From Tipo Where TipCodigo = " & Val(sFind)
        Else
            Cons = "Select TipCodigo, TipAbreviacion, TipNombre as 'Nombre' From Tipo Where TipNombre Like " & prj_SetFilterFind(sFind)
        End If
    Else
        sAux = "a Marca"
        If IsNumeric(sFind) Then
            Cons = "Select  MarCodigo, MarNombre as 'Nombre' From Marca Where MarCodigo = " & Val(sFind)
        Else
            Cons = "Select  MarCodigo, MarNombre as 'Nombre' From Marca Where MarNombre Like " & prj_SetFilterFind(sFind)
        End If
    End If
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        RsAux.MoveNext
        If Not RsAux.EOF Then
            RsAux.Close
            'Lista de ayuda
            If textB.Name = "tTipo" Then
                db_ShowHelpList Cons, textB, "Tipos de Artículos", 2
            Else
                db_ShowHelpList Cons, textB, "Marcas de Artículos", 1
            End If
        Else
            RsAux.MoveFirst
            With textB
                If textB.Name = "tTipo" Then
                    .Text = Trim(RsAux(2))
                Else
                    .Text = Trim(RsAux(1))
                End If
                .Tag = RsAux(0)
            End With
            If textB.Name = "tTipo" Then sAbrevTipo = Trim(RsAux(1))
        End If
    Else
        If MsgBox("No existe un" & sAux & " con el nombre ingresado." & vbCrLf & "¿Desea darle de alta?", vbQuestion + vbYesNo + vbDefaultButton2, "Nuevo") = vbYes Then
            'Llamo al mnt.
            
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
errFT:
    clsGeneral.OcurrioError "Error al buscar.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub cIVA_GotFocus()
    With cIVA
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub cIVA_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
End Sub

Private Sub Form_Load()
'
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    modProject.db_CloseConnect
End Sub

Private Sub Label1_Click()
    prj_SetFocus tTipo
End Sub

Private Sub Label2_Click()
    prj_SetFocus tMarca
End Sub

Private Sub Label3_Click()
    prj_SetFocus tDescripcion
End Sub

Private Sub Label4_Click()
    prj_SetFocus tNombre
End Sub

Private Sub Label5_Click()
    prj_SetFocus tProveedor
End Sub

Private Sub Label6_Click()
    prj_SetFocus tOrigen
End Sub

Private Sub Label7_Click()
    prj_SetFocus cIVA
End Sub

Private Sub Label8_Click()
    prj_SetFocus tCodigo
End Sub

Private Sub lParecido_Click()
    prj_SetFocus tParecidoA
End Sub

Private Sub tCodigo_GotFocus()
    With tCodigo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If prmID = 0 Then
            '........................................................Nuevo
            If IsNumeric(tCodigo.Text) Then
                If Val(tTipo.Tag) = 0 Then
                    prj_SetFocus tTipo
                ElseIf Val(tMarca.Tag) = 0 Then
                    prj_SetFocus tMarca
                Else
                    prj_SetFocus tDescripcion
                End If
            End If
            '........................................................Nuevo
        Else
            '........................................................Edición
            
            '........................................................Edición
            prj_SetFocus tTipo
        End If
    End If
End Sub

Private Sub tDescripcion_GotFocus()
    With tDescripcion
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tDescripcion_KeyPress(KeyAscii As Integer)
On Error GoTo errKP
Dim sName As String

    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    If tTipo.Tag = "" Then prj_SetFocus tTipo: Exit Sub
    If tMarca.Tag = "" Then prj_SetFocus tMarca: Exit Sub
    
    If UCase(Trim(tTipo.Text)) <> "N/D" And sAbrevTipo <> "" Then sName = sAbrevTipo & " "
    If UCase(Trim(tMarca.Text)) <> "N/D" Then sName = sName & Trim(tMarca.Text) & " "
    If Trim(sName) = "" Then sName = ""
    If UCase(Trim(tDescripcion.Text)) <> "N/D" Then sName = sName & Trim(tDescripcion)
    
    If Trim(tNombre.Text) <> "" And Trim(tNombre.Text) <> Trim(sName) Then
        If MsgBox("Desea modificar el nombre del artículo.", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
    End If
    tNombre.Text = Trim(sName)
    prj_SetFocus tNombre
    Exit Sub
errKP:
    clsGeneral.OcurrioError "Error al formar el nombre del artículo.", Err.Description
End Sub

Private Sub tMarca_Change()
    tMarca.Tag = ""
End Sub

Private Sub tMarca_GotFocus()
    With tMarca
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyReturn
              If Val(tMarca.Tag) = 0 Then db_FindTipoMarca tMarca.Text, tMarca
                If Val(tMarca.Tag) > 0 Then prj_SetFocus tDescripcion
            Case vbKeyF2
        End Select
    End If
End Sub

Private Sub tNombre_GotFocus()
    With tNombre
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then prj_SetFocus tProveedor
End Sub

Private Sub tOrigen_Change()
    tOrigen.Tag = ""
End Sub

Private Sub tOrigen_GotFocus()
    With tOrigen
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tOrigen_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tOrigen.Text) > 0 Then
            prj_SetFocus cIVA
        Else
        End If
    End If
End Sub

Private Sub tParecidoA_Change()
    tParecidoA.Tag = ""
End Sub

Private Sub tParecidoA_GotFocus()
    With tParecidoA
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tParecidoA_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
    End If
End Sub

Private Sub tProveedor_Change()
    tProveedor.Tag = ""
End Sub

Private Sub tProveedor_GotFocus()
    With tProveedor
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tProveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tProveedor.Tag) > 0 Then
            prj_SetFocus tOrigen
        Else
            'Busco.
        End If
    End If
End Sub

Private Sub tTipo_Change()
    tTipo.Tag = ""
    sAbrevTipo = ""
End Sub

Private Sub tTipo_GotFocus()
    With tTipo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyReturn
                If Val(tTipo.Tag) = 0 Then db_FindTipoMarca tTipo.Text, tTipo
                If Val(tTipo.Tag) > 0 Then prj_SetFocus tMarca
            Case vbKeyF2
                'Edición
        End Select
    End If
End Sub
