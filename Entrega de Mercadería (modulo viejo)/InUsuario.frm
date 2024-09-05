VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form InUsuario 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Abrir nueva sección"
   ClientHeight    =   1980
   ClientLeft      =   3900
   ClientTop       =   3435
   ClientWidth     =   4275
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "InUsuario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   741
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "grabar"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.TextBox tTecla 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox tDigito 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox tNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "(F1 a F12 - excepto F10)"
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Tecla abreviada:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese los datos para identificar al nuevo usuario."
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3600
      Picture         =   "InUsuario.frx":0442
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Nombre:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Dígito de Firma:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   -120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "InUsuario.frx":0884
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "InUsuario.frx":0B9E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuGrabar 
         Caption         =   "&Grabar"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "&Del formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "InUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsAux2 As rdoResultset

Dim gUsuarioCodigo As Long
Dim gUsuarioNombre As String
Dim gUsuarioTecla As Long

'Propiedades: Codigo y Nombre de Ususario...---------------
Public Property Get pUsuarioCodigo() As Long
    pUsuarioCodigo = gUsuarioCodigo
End Property
Public Property Let pUsuarioCodigo(Codigo As Long)
    gUsuarioCodigo = Codigo
End Property

Public Property Get pUsuarioNombre() As String
    pUsuarioNombre = gUsuarioNombre
End Property
Public Property Let pUsuarioNombre(Texto As String)
    gUsuarioNombre = Texto
End Property

Public Property Get pUsuarioTecla() As Long
    pUsuarioTecla = gUsuarioTecla
End Property
Public Property Let pUsuarioTecla(Codigo As Long)
    gUsuarioTecla = Codigo
End Property

'---------------------------------------------------------------------

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Call MnuVolver_Click
End Sub

Private Sub Form_Load()
    tTecla.Tag = "0"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    Forms(Forms.Count - 2).SetFocus
    
End Sub

Private Sub Label1_Click()
    Foco tDigito
End Sub

Private Sub Label2_Click()
    Foco tNombre
End Sub

Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuVolver_Click()
    gUsuarioCodigo = 0
    gUsuarioNombre = ""
    Unload Me
End Sub

Private Sub tDigito_GotFocus()
    tDigito.SelStart = 0
    tDigito.SelLength = Len(tDigito.Text)
End Sub

Private Sub tDigito_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn And IsNumeric(tDigito.Text) Then
        Cons = "Select * From Usuario Where UsuDigito = " & tDigito.Text
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly)
        
        If RsAux.EOF Then
            tDigito.Tag = ""
            MsgBox "No existe un usuario para el dígito ingresado.", vbExclamation, "ATENCIÓN"
        Else
            tDigito.Tag = RsAux!UsuCodigo
            tNombre.Text = Trim(RsAux!UsuIdentificacion)
            Foco tNombre
        End If
        RsAux.Close
    End If
    
End Sub

Private Sub tNombre_GotFocus()
    tNombre.SelStart = 0
    tNombre.SelLength = Len(tNombre.Text)
End Sub

Private Sub tNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(tNombre.Text) <> "" Then Foco tTecla
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        
        Case "grabar": AccionGrabar
        Case "salir": Call MnuVolver_Click
            
    End Select

End Sub

Private Sub AccionGrabar()

    If Trim(tNombre.Text) = "" Or Not IsNumeric(tDigito.Text) Then
        MsgBox "Los datos ingresados no son correctos.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
        
    If tDigito.Tag = "" Then tDigito.Tag = BuscoUsuario(CInt(tDigito.Text))
    
    If CInt(tDigito.Tag) = 0 Then Exit Sub
    
    If Trim(tTecla.Tag) = "0" Then
        MsgBox "Ingrese la tecla de función abreviada para llamar al formulario.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    If MsgBox("Su tecla de acceso directo a la sección es " & Trim(tTecla.Text) & Chr(vbKeyReturn) & "Confirma abrir la nueva sección.", vbQuestion + vbYesNo, "Nueva Sección") = vbNo Then Exit Sub
    
    On Error GoTo errGrabar
    gUsuarioCodigo = tDigito.Tag
    gUsuarioNombre = Trim(tNombre.Text)
    gUsuarioTecla = CLng(tTecla.Tag)
    Unload Me
    Exit Sub

errGrabar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al asignar las propiedades de la sección."
End Sub

Private Sub tTecla_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF1: tTecla.Tag = vbKeyF1: tTecla.Text = "F1"
        Case vbKeyF2: tTecla.Tag = vbKeyF2: tTecla.Text = "F2"
        Case vbKeyF3: tTecla.Tag = vbKeyF3: tTecla.Text = "F3"
        Case vbKeyF4: tTecla.Tag = vbKeyF4: tTecla.Text = "F4"
        Case vbKeyF5: tTecla.Tag = vbKeyF5: tTecla.Text = "F5"
        Case vbKeyF6: tTecla.Tag = vbKeyF6: tTecla.Text = "F6"
        Case vbKeyF7: tTecla.Tag = vbKeyF7: tTecla.Text = "F7"
        Case vbKeyF8: tTecla.Tag = vbKeyF8: tTecla.Text = "F8"
        Case vbKeyF9: tTecla.Tag = vbKeyF9: tTecla.Text = "F9"
        Case vbKeyF11: tTecla.Tag = vbKeyF11: tTecla.Text = "F11"
        Case vbKeyF12: tTecla.Tag = vbKeyF12: tTecla.Text = "F12"
    End Select
    
End Sub

Private Sub tTecla_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And tTecla.Tag <> "0" Then AccionGrabar
End Sub
