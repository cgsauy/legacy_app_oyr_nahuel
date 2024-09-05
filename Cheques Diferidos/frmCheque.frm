VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCheque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Cheques Diferidos"
   ClientHeight    =   4590
   ClientLeft      =   1545
   ClientTop       =   2085
   ClientWidth     =   7845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCheque.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7845
   Begin VB.TextBox tUsuario 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7200
      MaxLength       =   50
      TabIndex        =   13
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox tComentario 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   11
      Top             =   4200
      Width           =   6975
   End
   Begin VB.PictureBox pCheque 
      Height          =   2480
      Left            =   120
      ScaleHeight     =   2415
      ScaleWidth      =   7515
      TabIndex        =   14
      Top             =   720
      Width           =   7575
      Begin VB.ComboBox cMoneda 
         Height          =   315
         Left            =   5160
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   345
         Width           =   735
      End
      Begin VB.TextBox tImporte 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5880
         TabIndex        =   3
         Text            =   "1,565.23"
         Top             =   345
         Width           =   1455
      End
      Begin VB.TextBox tNumero 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         MaxLength       =   7
         TabIndex        =   1
         Text            =   "371101"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox tSerie 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "04"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox tFPago 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3720
         TabIndex        =   5
         Text            =   "1,565.23"
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox tFEmision 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   840
         TabIndex        =   6
         Text            =   "1,565.23"
         Top             =   2055
         Width           =   3375
      End
      Begin MSMask.MaskEdBox tBanco 
         Height          =   315
         Left            =   6600
         TabIndex        =   4
         Top             =   720
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   327681
         BackColor       =   16777215
         ForeColor       =   0
         PromptInclude   =   0   'False
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99-999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SERIE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "NUMERO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   25
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "C H E Q U E        D E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   24
         Top             =   165
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "PAGO DIFERIDO"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   23
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "PAGUESE DESDE EL:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   22
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "LA SUMA DE:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lImporte 
         BackStyle       =   0  'Transparent
         Caption         =   "Cinco mil trescientos cincuenta y ocho"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   1485
         Width           =   7215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "EMISIÓN:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2115
         Width           =   855
      End
      Begin VB.Label lBanco 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PAGO DIFERIDO"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   6015
      End
      Begin VB.Image Image2 
         Height          =   465
         Left            =   120
         Picture         =   "frmCheque.frx":0442
         Top             =   120
         Width           =   12000
      End
      Begin VB.Image Image3 
         Height          =   465
         Left            =   120
         Picture         =   "frmCheque.frx":187C
         Top             =   2145
         Width           =   12000
      End
      Begin VB.Image Image1 
         Height          =   2985
         Left            =   0
         Picture         =   "frmCheque.frx":2CB6
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7560
      End
   End
   Begin MSMask.MaskEdBox tCiF 
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   375
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      ForeColor       =   12582912
      MaxLength       =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#.###.###-#"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox tRucF 
      Height          =   285
      Left            =   1305
      TabIndex        =   19
      Top             =   375
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      ForeColor       =   12582912
      PromptInclude   =   0   'False
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99 999 999 9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox tCiC 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   3525
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      ForeColor       =   12582912
      MaxLength       =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#.###.###-#"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox tRucC 
      Height          =   285
      Left            =   1305
      TabIndex        =   9
      Top             =   3525
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   503
      _Version        =   327681
      Appearance      =   0
      ForeColor       =   12582912
      PromptInclude   =   0   'False
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99 999 999 9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label9 
      Caption         =   "&Dígito:"
      Height          =   255
      Left            =   7200
      TabIndex        =   12
      Top             =   3960
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000006&
      BorderWidth     =   2
      X1              =   0
      X2              =   7800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lClienteC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2760
      TabIndex        =   28
      Top             =   3525
      Width           =   4935
   End
   Begin VB.Label Label12 
      Caption         =   "Co&mentarios:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Cliente del &Cheque:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3285
      Width           =   1575
   End
   Begin VB.Label lClienteF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2760
      TabIndex        =   27
      Top             =   375
      Width           =   4935
   End
   Begin VB.Label Label5 
      Caption         =   "Cliente &Factura:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   135
      Width           =   1215
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuOpL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGrabar 
         Caption         =   "&Grabar"
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "&Cancelar"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuSaForm 
         Caption         =   "Del formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gCheque As Long, gTipoLlamado As Integer
Dim sNuevo As Boolean

Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then If cMoneda.ListIndex <> -1 Then Foco tImporte
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Height = 5280
    
    LimpioCampos
    lClienteF.Tag = 0: lClienteC.Tag = 0

    Cons = "Select MonCodigo, MonSigno from Moneda"
    CargoCombo Cons, cMoneda, ""
    
    If gTipoLlamado = TipoLlamado.IngresoNuevo Then AccionNuevo Else: EstadoCampos False
    
End Sub

Private Sub AccionNuevo()

    On Error Resume Next
    sNuevo = True
    EstadoCampos True
    
    MnuNuevo.Enabled = False
    MnuGrabar.Enabled = True
    MnuCancelar.Enabled = True
    
    Foco tCiF
    
End Sub

Private Sub AccionCancelar()

    sNuevo = False
    EstadoCampos False
    LimpioCampos
    
    MnuNuevo.Enabled = True
    MnuGrabar.Enabled = False
    MnuCancelar.Enabled = False
    
    tCiF.Text = FormatoCedula: tCiC.Text = FormatoCedula
    tRucF.Text = "": tRucC.Text = ""
    lClienteF.Caption = "": lClienteC.Caption = ""
    lClienteF.Tag = 0: lClienteC.Tag = 0
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If sNuevo Then
        If MsgBox("Ud. no ha completado el ingreso del cheque." & Chr(vbKeyReturn) & "Desea salir del formulario sin grabar.", vbExclamation + vbOKCancel + vbDefaultButton2, "ATENCIÓN") = vbCancel Then
            Cancel = 1
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Forms(Forms.Count - 2).SetFocus
End Sub

Private Sub Label9_Click()
    Foco tUsuario
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuNuevo_Click()
    AccionNuevo
End Sub

Private Sub MnuSaForm_Click()
    Unload Me
End Sub

Private Sub tBanco_Change()
    tBanco.Tag = ""
End Sub

Private Sub tBanco_GotFocus()
    tBanco.SelStart = 0
    tBanco.SelLength = 6
End Sub

Private Sub tBanco_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn And Len(tBanco.Text) = 5 Then
        If BuscoBancoEmisor(tBanco.Text) Then Foco tFPago
    End If

End Sub

Private Sub tCiC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        On Error Resume Next
        tRucC.Text = ""
        AccionBuscarClientes tCiC, lClienteC, TipoCliente.Persona
        If Val(lClienteC.Tag) <> 0 Then Foco tComentario
    End If
End Sub

Private Sub tCiC_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Trim(tCiC.Text) = FormatoCedula Then tRucC.SetFocus: Exit Sub
        tRucC.Text = ""
        
        Screen.MousePointer = 11
        If Len(clsGeneral.QuitoFormatoCedula(tCiC.Text)) = 7 Then tCiC.Text = clsGeneral.AgregoDigitoControlCI(tCiC.Text)
            
        If BuscoClienteCIRUC(clsGeneral.QuitoFormatoCedula(tCiC.Text), lClienteC, CI:=True) Then Foco tComentario
        Screen.MousePointer = 0
    End If

End Sub

Private Sub tCiF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        On Error Resume Next
        tRucF.Text = ""
        AccionBuscarClientes tCiF, lClienteF, TipoCliente.Persona
        If Val(lClienteF.Tag) <> 0 Then Foco tSerie
    End If
End Sub

Private Sub tCiF_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tCiF.Text) = FormatoCedula Then tRucF.SetFocus: Exit Sub
        tRucF.Text = ""
        
        Screen.MousePointer = 11
        If Len(clsGeneral.QuitoFormatoCedula(tCiF.Text)) = 7 Then tCiF.Text = clsGeneral.AgregoDigitoControlCI(tCiF.Text)
            
        If BuscoClienteCIRUC(clsGeneral.QuitoFormatoCedula(tCiF.Text), lClienteF, CI:=True) Then Foco tSerie
        Screen.MousePointer = 0
    End If

End Sub

Private Function BuscoClienteCIRUC(CiRuc As String, aControl As Control, Optional CI As Boolean = False, Optional RUC As Boolean = False) As Boolean

    On Error GoTo errBuscar
    BuscoClienteCIRUC = False
    
    If CI Then
        'Valido la Cédula ingresada----------
        If Len(CiRuc) <> 8 Then
            Screen.MousePointer = 0
            MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
            Exit Function
        End If
        If Not clsGeneral.CedulaValida(CiRuc) Then
            Screen.MousePointer = 0
            MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
            Exit Function
        End If
    End If
    
    Screen.MousePointer = 11
    Cons = "Select CliCodigo, CliCiRuc, CliTipo, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2) " _
           & " From Cliente, CPersona " _
           & " Where CliCiRuc = '" & CiRuc & "'" _
           & " And CliCodigo = CPeCliente " _
                                                & " UNION " _
           & " Select CliCodigo, CliCiRuc, CliTipo, Nombre = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')')  " _
           & " From Cliente, CEmpresa " _
           & " Where CliCiRuc = '" & CiRuc & "'" _
           & " And CliCodigo = CEmCliente"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not RsAux.EOF Then
        aControl.Caption = " " & Trim(RsAux!Nombre)
        aControl.Tag = RsAux!CliCodigo
        BuscoClienteCIRUC = True
    
    Else
        aControl.Caption = ""
        aControl.Tag = 0
        Screen.MousePointer = 0
        MsgBox "No existe un cliente para los datos ingresados ingresados.", vbExclamation, "ATENCIÓN"
    End If
    
    RsAux.Close
    Screen.MousePointer = 0
    Exit Function

errBuscar:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el cliente." & Err.Description
End Function

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tUsuario
End Sub

Private Sub tFEmision_Change()
    tFEmision.Tag = ""
End Sub

Private Sub tFEmision_GotFocus()
    tFEmision.Text = tFEmision.Tag
    tFEmision.SelStart = 0
    tFEmision.SelLength = Len(tFEmision.Text)
End Sub

Private Sub tFEmision_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Trim(tFEmision.Tag) = "" Then
             If IsDate(tFEmision.Text) Then
                aTexto = tFEmision.Text
                tFEmision.Text = FormateoFecha(tFEmision.Text)
                tFEmision.Tag = Format(aTexto, "dd/mm/yyyy")
                Foco tCiC
            End If
        Else
            Foco tCiC
        End If
    End If

End Sub

Private Sub tFPago_Change()
    tFPago.Tag = ""
End Sub

Private Sub tFPago_GotFocus()
    tFPago.Text = tFPago.Tag
    tFPago.SelStart = 0
    tFPago.SelLength = Len(tFPago.Text)
End Sub

Private Sub tFPago_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tFPago.Tag) = "" Then
             If IsDate(tFPago.Text) Then
                aTexto = tFPago.Text
                tFPago.Text = FormateoFecha(tFPago.Text)
                tFPago.Tag = Format(aTexto, "dd/mm/yyyy")
                Foco tFEmision
            End If
        Else
            Foco tFEmision
        End If
    End If
    
End Sub

Private Sub tImporte_GotFocus()
    tImporte.SelStart = 0
    tImporte.SelLength = Len(tImporte.Text)
End Sub

Private Sub tImporte_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And Trim(tImporte.Text) <> "" Then
        If Not IsNumeric(tImporte.Text) Then MsgBox "El importe ingresado no es correcto. Verifique los datos.", vbExclamation, "ATENCIÓN": Exit Sub
        tImporte.Text = Format(tImporte.Text, FormatoMonedaP)
        lImporte.Caption = ImporteATexto(tImporte.Text) & Space(400 - Len(Trim(lImporte.Caption)))
        Foco tBanco
    End If
        
End Sub

Private Sub tNumero_GotFocus()
    tNumero.SelStart = 0
    tNumero.SelLength = Len(tNumero.Text)
End Sub

Private Sub tNumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(tNumero.Text) <> "" Then Foco cMoneda
End Sub

Private Sub tRucC_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        On Error Resume Next
        tCiC.Text = FormatoCedula
        AccionBuscarClientes tRucC, lClienteC, TipoCliente.Empresa
        If Val(lClienteC.Tag) <> 0 Then Foco tComentario
    End If
    
End Sub

Private Sub tRucC_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tRucC.Text) = "" Then tCiC.SetFocus: Exit Sub
        tCiC.Text = FormatoCedula
        
        Screen.MousePointer = 11
        If BuscoClienteCIRUC(Trim(tRucC.Text), lClienteC, RUC:=True) Then Foco tComentario
        
        Screen.MousePointer = 0
    End If

End Sub

Private Sub tRucF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        On Error Resume Next
        tCiF.Text = FormatoCedula
        AccionBuscarClientes tRucF, lClienteF, TipoCliente.Empresa
        If Val(lClienteF.Tag) <> 0 Then Foco tSerie
    End If
End Sub

Private Sub tRucF_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tRucF.Text) = "" Then tCiF.SetFocus: Exit Sub
        tCiF.Text = FormatoCedula
        
        Screen.MousePointer = 11
        If BuscoClienteCIRUC(Trim(tRucF.Text), lClienteF, RUC:=True) Then Foco tSerie
        
        Screen.MousePointer = 0
    End If
    
End Sub

Private Sub tSerie_GotFocus()
    tSerie.SelStart = 0
    tSerie.SelLength = Len(tSerie.Text)
End Sub

Private Sub tSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(tSerie.Text) <> "" Then Foco tNumero
End Sub

Private Function BuscoBancoEmisor(Codigo As String) As Boolean
    
Dim Banco As String, Sucursal As String

    On Error GoTo errCargar
    Banco = Mid(Codigo, 1, 2)
    Sucursal = Mid(Codigo, 3, 3)
    BuscoBancoEmisor = True
    
    Cons = "Select * from BancoSSFF, SucursalDeBanco" _
          & "  Where BanCodigoB = " & Banco _
          & "  And SBaCodigoS = " & Sucursal _
          & "  And SBaBanco = BanCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        lBanco.Caption = UCase(Trim(RsAux!BanNombre) & " - " & Trim(RsAux!SBaNombre))
        tBanco.Tag = RsAux!BanCodigo & "-" & RsAux!SBaCodigo
    Else
        MsgBox "No existe un banco para el código ingresado.", vbExclamation, "ATENCIÓN"
        tBanco.Tag = ""
        BuscoBancoEmisor = False
    End If
    RsAux.Close
    Exit Function

errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al cargar el banco emisor."
End Function

Private Function FormateoFecha(fecha As String) As String
    FormateoFecha = Format(fecha, "d") & " de " & Format(fecha, "Mmmm") & " de " & Format(fecha, "yyyy")
End Function

Private Sub AccionGrabar()

    If Not ValidoDatos Then Exit Sub
    
    On Error GoTo errorBT
    Screen.MousePointer = 11
    'Valido que no exista un cheque con las mismas caracteristicas---------------------------------------------------------
    'Hay que validar que el cheque no este ingresado (Banco, Sucursal, Numero)
    Cons = "Select * from ChequeDiferido " _
           & " Where CDiBanco = " & ZBancoSucursal(tBanco.Tag, Banco:=True) _
           & " And CDiSucursal = " & ZBancoSucursal(tBanco.Tag, Sucursal:=True) _
           & " And CDiSerie = '" & Trim(tSerie.Text) & "'" _
           & " And CDiNumero = " & tNumero.Text
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        Screen.MousePointer = 0
        MsgBox "El cheque ya ha sido ingresado al sistema." & Chr(vbKeyReturn) _
                & "Verifique los datos ingresados en Seguimiento de Cheques.", vbCritical, "ATENCIÓN"
        RsAux.Close
        Exit Sub
    End If
    RsAux.Close
    '--------------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 0
    If MsgBox("Confirma grabar los datos del cheque.", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    FechaDelServidor
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET

    Cons = "Select * from ChequeDiferido Where CDiCodigo = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    RsAux.AddNew
    RsAux!CDiBanco = ZBancoSucursal(tBanco.Tag, Banco:=True)
    RsAux!CDiSucursal = ZBancoSucursal(tBanco.Tag, Sucursal:=True)
    RsAux!CDiSerie = Trim(tSerie.Text)
    RsAux!CDiNumero = tNumero.Text
    RsAux!CDiMoneda = cMoneda.ItemData(cMoneda.ListIndex)
    RsAux!CDiImporte = CCur(tImporte.Text)
    RsAux!CDiVencimiento = Format(tFPago.Tag, FormatoFH)
    RsAux!CDiLibrado = Format(tFEmision.Tag, FormatoFH)
    RsAux!CDiCliente = lClienteC.Tag
    RsAux!CDiUsuario = tUsuario.Tag
    RsAux!CDiIngresado = Format(gFechaServidor, FormatoFH)
    RsAux!CDiClienteFactura = lClienteF.Tag
    If Trim(tComentario.Text) <> "" Then RsAux!CDiComentario = Trim(tComentario.Text)
    RsAux.Update
    
    'Hago Movimiento de Caja  (Salida -- xq' ingreso un cheque).
    aTexto = "Dif. " & tSerie.Text & tNumero.Text & " del " & "(" & tBanco.FormattedText & ") " & Trim(lBanco.Caption)
    MovimientoDeCaja paMCChequeDiferido, gFechaServidor, paDisponibilidad, cMoneda.ItemData(cMoneda.ListIndex), CCur(tImporte.Text), aTexto, True
                                
    cBase.CommitTrans    'Fin de la TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    Screen.MousePointer = 0
    LimpioCampos
    
    If gTipoLlamado = TipoLlamado.IngresoNuevo Then
        If MsgBox("Desea ingresar un nuevo cheque.", vbQuestion + vbYesNo, "NUEVO CHEQUE") = vbYes Then
            Foco tSerie
        Else
            sNuevo = False: Unload Me
        End If
    Else
        AccionCancelar
    End If
    Exit Sub
    
errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación."
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación."
End Sub

Private Function ValidoDatos() As Boolean
    ValidoDatos = False
    On Error GoTo errValido
    
    If Val(tUsuario.Tag) = 0 Then
        MsgBox "Ingrese su dígito de usuario.", vbExclamation, "ATENCIÓN"
        Foco tUsuario: Exit Function
    End If
    
    If Val(lClienteF.Tag) = 0 Then
        MsgBox "Se debe ingresar el cliente titular de la operación.", vbExclamation, "ATENCIÓN"
        tCiF.SetFocus: Exit Function
    End If
    
    If Val(lClienteC.Tag) = 0 Then
        MsgBox "Se debe ingresar el cliente que emite el cheque.", vbExclamation, "ATENCIÓN"
        tCiC.SetFocus: Exit Function
    End If
    
    If Trim(tSerie.Text) = "" Then
        MsgBox "La serie del cheque no es correcta. Verifique", vbExclamation, "ATENCIÓN"
        Foco tSerie: Exit Function
    End If
    
    If Not IsNumeric(tNumero.Text) Then
        MsgBox "El número del cheque no es correcto. Verifique", vbExclamation, "ATENCIÓN"
        Foco tNumero: Exit Function
    End If
    
    If Not IsNumeric(tImporte.Text) Then
        MsgBox "El importe del cheque no es correcto. Verifique", vbExclamation, "ATENCIÓN"
        Foco tImporte: Exit Function
    End If
    
    If tBanco.Tag = "" Then
        MsgBox "El banco seleccionado no es correcto. Verifique", vbExclamation, "ATENCIÓN"
        Foco tBanco: Exit Function
    End If
    
    If Not IsDate(tFPago.Tag) Then
        MsgBox "La fecha de vencimiento ingresada no es correcta. Verifique", vbExclamation, "ATENCIÓN"
        Foco tFPago: Exit Function
    End If
    
    If Not IsDate(tFEmision.Tag) Then
        MsgBox "La fecha de emisión del cheque no es correcta. Verifique", vbExclamation, "ATENCIÓN"
        Foco tFEmision: Exit Function
    End If
    
    If CDate(tFPago.Tag) < CDate(tFEmision.Tag) Then
        MsgBox "La fecha de emisión del cheque no debe ser mayor a la del pago. Verifique", vbExclamation, "ATENCIÓN"
        Foco tFEmision: Exit Function
    End If
    
    'Valido que sea diferido
    FechaDelServidor
    If Format(tFPago.Tag, "dd/mm/yyyy") = Format(gFechaServidor, "dd/mm/yyyy") Then
        MsgBox "El cheque ingresado NO ES DIFERIDO, la fecha del vencimiento es la del día de hoy, por lo tanto ....  es un CHEQUE AL DÍA.", vbExclamation, "ATENCIÓN"
        Foco tFPago: Exit Function
    End If

    If Format(tFPago.Tag, "yyyy/mm/dd") < Format(gFechaServidor, "yyyy/mm/dd") Then
        MsgBox "La fecha de pago del cheque es menor al día de hoy. Verifique", vbExclamation, "ATENCIÓN"
        Foco tFPago: Exit Function
    End If
    ValidoDatos = True
    Exit Function

errValido:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al validar los datos. " & Err.Description
End Function

Private Sub CargoDatosCliente(Codigo As Long)

    On Error Resume Next
    lClienteF.Caption = ""
    Cons = "Select * from Cliente, CEmpresa Where CliCodigo = " & Codigo & " And CliCodigo = CEmCliente"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!CEmNombre) Then lClienteF.Caption = " " & UCase(Trim(RsAux!CEmNombre)) Else lClienteF.Caption = UCase(Trim(RsAux!CEmFantasia))
        tRucF.Text = Trim(RsAux!CliCIRuc)
        lClienteF.Tag = Codigo
    End If
    RsAux.Close
    
    If Trim(lClienteF.Caption) <> "" Then Exit Sub
    
    Cons = "Select * from Cliente, CPersona Where CliCodigo = " & Codigo & " And CliCodigo = CPeCliente"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        lClienteF.Caption = " " & ArmoNombre(RsAux!CPeApellido1, Format(RsAux!CPeApellido2, "#"), RsAux!CPeNombre1, Format(RsAux!CPeNombre2, "#"))
        tCiF.Text = RetornoFormatoCedula(RsAux!CliCIRuc)
        lClienteF.Tag = Codigo
    End If
    RsAux.Close
    
End Sub

Private Sub AccionBuscarClientes(txtCiRuc As Control, lblNombre As Control, aTipoCliente As Integer)
    
    Screen.MousePointer = 11
    On Error GoTo errCargar
    Dim aIdCliente As Long, aTipo As Integer
    Dim objBuscar As New clsBuscarCliente
    
    If aTipoCliente = TipoCliente.Persona Then objBuscar.ActivoFormularioBuscarClientes strConeccion:=txtConexion, Persona:=True
    If aTipoCliente = TipoCliente.Empresa Then objBuscar.ActivoFormularioBuscarClientes strConeccion:=txtConexion, Empresa:=True
    Me.Refresh
    
    aIdCliente = objBuscar.BCClienteSeleccionado
    aTipo = objBuscar.BCTipoClienteSeleccionado
    Set objBuscar = Nothing
    
    If aIdCliente <> 0 Then
        Dim aCodigo As Long
        aCodigo = aIdCliente
        
        On Error Resume Next
        lblNombre.Caption = ""
        
        If aTipo = TipoCliente.Empresa Then
            Cons = "Select * from Cliente, CEmpresa Where CliCodigo = " & aCodigo & " And CliCodigo = CEmCliente"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux.EOF Then
                If Not IsNull(RsAux!CEmNombre) Then lblNombre.Caption = " " & Trim(RsAux!CEmFantasia) & " (" & Trim(RsAux!CEmNombre) & ")" Else lClienteF.Caption = UCase(Trim(RsAux!CEmFantasia))
                If Not IsNull(RsAux!CliCIRuc) Then txtCiRuc.Text = Trim(RsAux!CliCIRuc) Else txtCiRuc.Text = ""
                lblNombre.Tag = aCodigo
            End If
            RsAux.Close
        End If
        
        If aTipo = TipoCliente.Persona Then
            Cons = "Select * from Cliente, CPersona Where CliCodigo = " & aCodigo & " And CliCodigo = CPeCliente"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux.EOF Then
                lblNombre.Caption = " " & ArmoNombre(RsAux!CPeApellido1, Format(RsAux!CPeApellido2, "#"), RsAux!CPeNombre1, Format(RsAux!CPeNombre2, "#"))
                If Not IsNull(RsAux!CliCIRuc) Then txtCiRuc.Text = RetornoFormatoCedula(RsAux!CliCIRuc) Else txtCiRuc.Text = FormatoCedula
                lblNombre.Tag = aCodigo
            End If
            RsAux.Close
        End If
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del cliente.", Err.Description
End Sub


Private Function ZBancoSucursal(Texto As String, Optional Banco As Boolean = False, Optional Sucursal As Boolean = False)
    If Banco Then ZBancoSucursal = Mid(Texto, 1, InStr(Texto, "-") - 1)
    If Sucursal Then ZBancoSucursal = Mid(Texto, InStr(Texto, "-") + 1, Len(Texto))
End Function

Public Property Get pCodigoMoneda() As Long: End Property
Public Property Let pCodigoMoneda(Numero As Long)
    On Error Resume Next
    BuscoCodigoEnCombo cMoneda, Numero
End Property

Public Property Get pCliente() As Long: End Property
Public Property Let pCliente(Numero As Long)
    On Error Resume Next
    CargoDatosCliente Numero
    Foco tSerie
End Property

Public Property Get pLlamado() As Integer: End Property
Public Property Let pLlamado(Numero As Integer)
    gTipoLlamado = Numero
End Property

Public Property Get pComentario() As String: End Property
Public Property Let pComentario(Texto As String)
    tComentario = Trim(Texto)
End Property


Private Sub EstadoCampos(Estado As Boolean)

    tCiF.Enabled = Estado
    tRucF.Enabled = Estado
    tCiC.Enabled = Estado
    tRucC.Enabled = Estado
    tComentario.Enabled = Estado
    
    tSerie.Enabled = Estado
    tNumero.Enabled = Estado
    tBanco.Enabled = Estado
    tImporte.Enabled = Estado
    tFPago.Enabled = Estado
    tFEmision.Enabled = Estado
    cMoneda.Enabled = Estado
    tUsuario.Enabled = Estado
    
    If Estado = False Then
        lClienteF.BackColor = Me.BackColor
        lClienteC.BackColor = Me.BackColor
        tComentario.BackColor = Me.BackColor
        tUsuario.BackColor = Me.BackColor
    Else
        lClienteF.BackColor = Blanco
        lClienteC.BackColor = Blanco
        tComentario.BackColor = Blanco
        tUsuario.BackColor = Blanco
    End If
    
End Sub

Private Sub LimpioCampos()

    tSerie.Text = "": tNumero.Text = ""
    tImporte.Text = ""
    tBanco.Text = "": lBanco.Caption = ""
    tFPago.Text = "": tFEmision.Text = ""
    lImporte.Caption = Space(400)
    
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        On Error Resume Next
        If IsNumeric(tUsuario.Text) Then tUsuario.Tag = BuscoUsuario(CInt(tUsuario.Text))
        If tUsuario.Tag <> 0 Then AccionGrabar
    End If
    
End Sub
