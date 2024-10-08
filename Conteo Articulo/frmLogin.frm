VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Clave"
   ClientHeight    =   1290
   ClientLeft      =   4065
   ClientTop       =   3045
   ClientWidth     =   3225
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   762.174
   ScaleMode       =   0  'User
   ScaleWidth      =   3028.101
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox tUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   305
      Left            =   120
      MaxLength       =   12
      TabIndex        =   1
      Top             =   330
      Width           =   1605
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   339
      Left            =   2160
      TabIndex        =   3
      Top             =   900
      Width           =   1012
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   339
      Left            =   2160
      TabIndex        =   4
      Top             =   480
      Width           =   1012
   End
   Begin VB.TextBox tClave 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   305
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   900
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "&Contraseña:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   690
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Identificación:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   165
      TabIndex        =   0
      Top             =   105
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gLoginOk As Boolean               'Logion exitoso

Public prmIDUsuario As Long
Public prmNombre As String

Private Sub cmdCancel_Click()
    gLoginOk = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    VerificoUsuario
End Sub

Private Sub Form_Activate()

    On Error Resume Next
    Me.Refresh
    tUsuario.SetFocus
    
End Sub

Private Sub Form_Load()
    
    prmIDUsuario = 0
    prmNombre = ""
    
    gLoginOk = False
    
End Sub

Private Sub VerificoUsuario()
    
    If Not ValidoIngreso Then Exit Sub
    
    On Error GoTo errBuscar
    Screen.MousePointer = 11
    
    Dim aUser As Long
    Cons = "SELECT UsuCodigo, rTrim(UsuIdentificacion) FROM Usuario" _
           & " Where UsuIdentificacion = '" & Trim(tUsuario) & "' AND UsuHabilitado=1"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        aUser = RsAux("UsuCodigo")
        prmNombre = Trim(RsAux(1))
    End If
    RsAux.Close
    
    If aUser = 0 Then
        prmNombre = ""
        MsgBox "Usuario incorrecto.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    If Not miConexion.ValidoClave(aUser, tClave.Text) Then
        MsgBox "No se encontró un usuario para la identificación ingresada o la clave no es correcta.", vbCritical, "ATENCIÓN"
        Exit Sub
    End If
    
    prmIDUsuario = aUser
    
    
'    'VERIFICO USUARIO Y PASSWORD------------------------------------------------------------------------------
'    Cons = "SELECT * FROM Usuario" _
'           & " Where UsuIdentificacion = '" & Trim(tUsuario) & "'" _
'           & " And (UsuContraseña = '" & EncryptoString(Trim(LCase(tClave.Text))) & "' OR UsuContraseña = '" & EncryptoString(UCase(Trim(tClave.Text))) & "')"
'    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'    If RsAux.EOF Then
'        Screen.MousePointer = 0
'        MsgBox "No se encontró un usuario para la identificación ingresada o la clave no es correcta.", vbCritical, "ATENCIÓN"
'        RsAux.Close: Exit Sub
'    End If
'
'    If Not RsAux!UsuHabilitado Then
'        Screen.MousePointer = 0
'        MsgBox "Su usuario no está habilitado para acceder al sistema.", vbCritical, "ATENCIÓN"
'        RsAux.Close: Exit Sub
'    End If
'
'    prmIDUsuario = RsAux!UsuCodigo
'    prmNombre = Trim(RsAux!UsuIdentificacion)
'    RsAux.Close
    
    gLoginOk = True
    Screen.MousePointer = 0
    Unload Me
        
    Exit Sub

errBuscar:
    clsGeneral.OcurrioError "Error al buscar el usuario.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If Not gLoginOk Then
        prmIDUsuario = 0
        prmNombre = ""
    End If
    
End Sub

Private Sub Label1_Click()
    tClave.SelStart = 0
    tClave.SelLength = Len(tClave.Text)
End Sub

Private Sub lblLabels_Click()
    tUsuario.SelStart = 0: tUsuario.SelLength = Len(tUsuario.Text)
End Sub

Private Sub tClave_Change()
    'tmLogin.Enabled = False: tmLogin.Enabled = True
End Sub

Private Sub tClave_GotFocus()
    tClave.SelStart = 0: tClave.SelLength = Len(tClave.Text)
End Sub

Private Sub tClave_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tClave.Text) <> "" And Trim(tUsuario.Text) <> "" Then VerificoUsuario Else: tUsuario.SetFocus
    End If
    
End Sub

Private Sub tUsuario_GotFocus()
    tUsuario.SelStart = 0: tUsuario.SelLength = Len(tUsuario.Text)
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And Trim(tUsuario) <> "" Then
        On Error Resume Next
        'Hay 2 opciones 1) ingreso del dígito, 2) ingreso del nombre
        If IsNumeric(tUsuario.Text) Then
            Screen.MousePointer = 11
            Cons = "Select * From Usuario Where UsuDigito = " & Val(tUsuario.Text)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then tUsuario.Text = Trim(RsAux!UsuIdentificacion)
            RsAux.Close
            Screen.MousePointer = 0
        End If
        Foco tClave
    End If
    
End Sub

Private Function ValidoIngreso() As Boolean

    ValidoIngreso = False
    
    If Trim(tUsuario) = "" Then
        MsgBox "Ingrese su identificación de usuario.", vbInformation, "ATENCIÓN"
        tUsuario.SetFocus: Exit Function
    End If
    
    If Trim(tClave.Text) = "" Then
        MsgBox "Ingrese su clave de ingreso.", vbInformation, "ATENCIÓN"
        tClave.SetFocus: Exit Function
    End If
    
    ValidoIngreso = True
    
End Function

