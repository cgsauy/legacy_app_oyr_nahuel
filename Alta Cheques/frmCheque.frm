VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCheque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Cheques"
   ClientHeight    =   4545
   ClientLeft      =   2730
   ClientTop       =   2835
   ClientWidth     =   7755
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
   ScaleHeight     =   4545
   ScaleWidth      =   7755
   Begin VB.ComboBox cTipo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   1695
   End
   Begin VB.TextBox tUsuario 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   7200
      MaxLength       =   50
      TabIndex        =   16
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox tComentario 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   14
      Top             =   4200
      Width           =   6975
   End
   Begin VB.PictureBox pCheque 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2480
      Left            =   120
      ScaleHeight     =   2445
      ScaleWidth      =   7545
      TabIndex        =   17
      Top             =   780
      Width           =   7575
      Begin VB.TextBox tDocumento 
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   5160
         MaxLength       =   12
         TabIndex        =   9
         Top             =   2040
         Width           =   915
      End
      Begin VB.ComboBox cMoneda 
         Height          =   315
         Left            =   5160
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   3
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
         TabIndex        =   2
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
         TabIndex        =   7
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
         TabIndex        =   8
         Text            =   "1,565.23"
         Top             =   2055
         Width           =   3375
      End
      Begin MSMask.MaskEdBox tBanco 
         Height          =   315
         Left            =   6480
         TabIndex        =   6
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "DOC:"
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
         Left            =   4740
         TabIndex        =   33
         Top             =   2115
         Width           =   495
      End
      Begin VB.Label lDoc 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   6120
         TabIndex        =   32
         Top             =   2100
         Width           =   1335
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
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lDiferido0 
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
         TabIndex        =   27
         Top             =   165
         Width           =   1575
      End
      Begin VB.Label lDiferido1 
         BackStyle       =   0  'Transparent
         Caption         =   "PAGO DIFERIDO"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   26
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lDiferido2 
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   21
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
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
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
      TabIndex        =   20
      Top             =   435
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      ForeColor       =   12582912
      PromptInclude   =   0   'False
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
      TabIndex        =   22
      Top             =   435
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   0
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
      TabIndex        =   11
      Top             =   3585
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      ForeColor       =   12582912
      PromptInclude   =   0   'False
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
      TabIndex        =   12
      Top             =   3585
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   0
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
   Begin VB.Label lIdCheque 
      BackStyle       =   0  'Transparent
      Caption         =   "PAGO DIFERIDO"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2340
      TabIndex        =   34
      Top             =   90
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "&TIPO DE CHEQUE"
      Height          =   255
      Left            =   4500
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "&Dígito:"
      Height          =   255
      Left            =   7200
      TabIndex        =   15
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2760
      TabIndex        =   31
      Top             =   3585
      Width           =   4935
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Co&mentarios:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente del &Cheque:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3345
      Width           =   1575
   End
   Begin VB.Label lClienteF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2760
      TabIndex        =   30
      Top             =   435
      Width           =   4935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente &Factura:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   195
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
   Begin VB.Menu MnuAyuda 
      Caption         =   "&?"
      Begin VB.Menu MnuHelp 
         Caption         =   "Ayuda"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public prmTipo As Integer                   'Tipo de Cheque a Ingresar
Public prmIDRecCtdo As Long                 'ID de rec o Ctdo asociado al cheque
Public prmImporte As Currency               'Importe Total del Cheque
Public prmTAG_LiqCamion As Integer          'Tag con el cual intento grabar los cheques para una liquidación.

Private bYaInserteChequeLC As Boolean

Dim gCheque As Long
Dim sNuevo As Boolean

Dim aTexto As String

Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then If cMoneda.ListIndex <> -1 Then Foco tImporte
End Sub

Private Sub cTipo_Click()
    Select Case cTipo.ListIndex
        Case 0
                lDiferido0.Visible = False: lDiferido1.Visible = False
                lDiferido2.Visible = False: tFPago.Visible = False
                
        Case 1
                lDiferido0.Visible = True: lDiferido1.Visible = True
                lDiferido2.Visible = True: tFPago.Visible = True
    End Select
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    FechaDelServidor
    
    Me.BackColor = RGB(250, 240, 230)
    Me.Top = (Screen.Height - Me.Height) / 2: Me.Left = (Screen.Width - Me.Width) / 2
    Me.Height = 5235
    
    LimpioCampos
    lClienteF.Tag = 0: lClienteC.Tag = 0

    Cons = "Select MonCodigo, MonSigno from Moneda Where MonFactura = 1 Order by MonSigno"
    CargoCombo Cons, cMoneda, ""
    
    cTipo.AddItem "AL DIA "
    cTipo.AddItem "DIFERIDO"
    
    If prmTipo <> -1 Then
        cTipo.ListIndex = prmTipo
        AccionNuevo
        CargoValoresParametros
    Else
        EstadoCampos False
        cTipo.ListIndex = 0
    End If
    
End Sub

Private Sub CargoValoresParametros()
On Error GoTo errBDoc

Dim adTexto As String, adCodigo As Long, adCliente As Long, adTextoFMT As String, adImporte As Currency
Dim adMoneda As Long, adTipo As Integer

    adMoneda = paMonedaPesos
    adImporte = prmImporte
    
    If prmIDRecCtdo <> 0 Then
        Cons = "Select * from Documento " & _
                   " Where DocCodigo = " & prmIDRecCtdo
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            adCodigo = prmIDRecCtdo
            adMoneda = RsAux!DocMoneda
            adTipo = RsAux!DocTipo
            adTexto = zDocumento(RsAux!DocTipo, RsAux!DocSerie, RsAux!DocNumero, adTextoFMT)
            adCliente = RsAux!DocCliente
            
            If prmImporte = 0 Then adImporte = Format(RsAux!DocTotal, "#,##0.00")
        End If
        RsAux.Close
    
    End If
    
    If adImporte <> 0 Then
        BuscoCodigoEnCombo cMoneda, adMoneda
        If cTipo.ListIndex = 0 Then tImporte.Text = Format(adImporte, "#,##0.00")
    End If
    
    If adCodigo = 0 Then
        If prmIDRecCtdo <> 0 Then lDoc.Caption = " No Existe !!" Else lDoc.Caption = ""
        lDoc.Tag = 0
        Exit Sub
    End If
    
    tDocumento.Text = adTextoFMT
    lDoc.Tag = adCodigo: lDoc.Caption = adTexto
        
    BuscoCodigoEnCombo cMoneda, adMoneda
    CargoDatosCliente adCliente
        
    tFEmision.Text = Format(gFechaServidor, "dd/mm/yyyy")
    Call tFEmision_KeyPress(vbKeyReturn)
    
    Select Case adTipo
        Case TipoDocumento.Contado
        
        Case TipoDocumento.ReciboDePago
            adTextoFMT = ""
            Cons = "Select DocSerie, DocNumero from DocumentoPago, Documento" & _
                        " Where DPaDocQSalda = " & adCodigo & _
                        " And DPaDocASaldar = DocCodigo" & _
                        " Group by DocSerie, DocNumero"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            Do While Not RsAux.EOF
                If Trim(adTextoFMT) <> "" Then adTextoFMT = adTextoFMT & "; "
                adTextoFMT = adTextoFMT & Trim(RsAux!DocSerie) & "-" & (RsAux!DocNumero)
                RsAux.MoveNext
            Loop
            RsAux.Close
            If Trim(adTextoFMT) <> "" Then
                If InStr(adTextoFMT, ";") <> 0 Then adTextoFMT = "Créditos " & adTextoFMT Else adTextoFMT = "Crédito " & adTextoFMT
            End If
             
            tComentario.Text = adTextoFMT
        
    End Select

    lIdCheque.Caption = " CHEQUE Nº 1"
    lIdCheque.Tag = 1: lIdCheque.Visible = True
    
    Exit Sub
    
errBDoc:
    clsGeneral.OcurrioError "Error al buscar el documento.", Err.Description
End Sub

Private Sub AccionNuevo()

    On Error Resume Next
    If cTipo.ListIndex = -1 Then
        MsgBox "Seleccione el tipo de cheque a ingresar." & vbCrLf & _
                    "Cheque Al Día o de Pago Diferido.", vbExclamation, "Falta Tipo de Cheque"
        cTipo.SetFocus: Exit Sub
    End If
    
    sNuevo = True
    EstadoCampos True
    
    MnuNuevo.Enabled = False
    MnuGrabar.Enabled = True
    MnuCancelar.Enabled = True
    
    Foco tCiF
    
End Sub

Private Sub AccionCancelar()
On Error Resume Next
    
    lIdCheque.Visible = False
    
    sNuevo = False
    EstadoCampos False
    LimpioCampos
    
    MnuNuevo.Enabled = True
    MnuGrabar.Enabled = False
    MnuCancelar.Enabled = False
    
    tCiF.Text = "": tCiC.Text = ""
    tRucF.Text = "": tRucC.Text = ""
    lClienteF.Caption = "": lClienteC.Caption = ""
    lClienteF.Tag = 0: lClienteC.Tag = 0
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If sNuevo Then
        If MsgBox("Ud. no ha completado el ingreso del cheque." & vbCrLf & _
                        "Desea salir del formulario sin grabar.", vbExclamation + vbYesNo + vbDefaultButton2, "Cheque Incompleto") = vbNo Then
            Cancel = 1
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EndMain
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

Private Sub MnuHelp_Click()
On Error GoTo errHelp
    Screen.MousePointer = 11
    
    Dim aFile As String
    Cons = "Select * from Aplicacion Where AplNombre = '" & prmKeyApp & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux!AplHelp) Then aFile = Trim(RsAux!AplHelp)
    RsAux.Close
    
    If aFile <> "" Then EjecutarApp aFile
    
    Screen.MousePointer = 0
    Exit Sub
    
errHelp:
    clsGeneral.OcurrioError "Error al activar el archivo de ayuda.", Err.Description
    Screen.MousePointer = 0
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
    tBanco.SelStart = 0: tBanco.SelLength = 6
End Sub

Private Sub tBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errBancos
    
    If KeyCode = vbKeyF1 Then
        Screen.MousePointer = 11
        Cons = "Select BanCodigoB as 'Nº B', BanNombre as 'Banco' " & _
                    " From BancoSSFF " & _
                    " Order by BanNombre"
                    
        Dim objHelp As New clsListadeAyuda, mSel As String
        mSel = ""
        If objHelp.ActivarAyuda(cBase, Cons, anchoform:=4500, Titulo:="Lista de Bancos") Then
            mSel = Format(objHelp.RetornoDatoSeleccionado(0), "00")
        End If
        Set objHelp = Nothing
        
        Me.Refresh
        Screen.MousePointer = 0
        If mSel <> "" Then
            tBanco.Text = mSel & "999"
            Call tBanco_KeyPress(vbKeyReturn)
        End If
    End If
    Exit Sub

errBancos:
    clsGeneral.OcurrioError "Error al buscar los bancos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tBanco_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn And Len(tBanco.Text) = 5 Then
        If BuscoBancoEmisor(tBanco.Text) Then
            If tFPago.Visible Then Foco tFPago Else Foco tFEmision
        End If
    End If

End Sub

Private Sub tCiC_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    Select Case KeyCode
        Case vbKeyF4
            tRucC.Text = ""
            AccionBuscarClientes tCiC, lClienteC, TipoCliente.Cliente
            If Val(lClienteC.Tag) <> 0 Then Foco tComentario
            
        Case vbKeyMultiply
            If Val(lClienteF.Tag) = 0 Then Exit Sub
            CargoDatosCliente Val(lClienteF.Tag), 2
            
            'If Trim(tCiF.Text) <> "" Then
            '    tCiC.Text = tCiF.Text
            '    tRucC.Text = ""
            '    BuscoClienteCIRUC Trim(tCiC.Text), lClienteC, CI:=True
            'End If
            
            'If Trim(tRucF.Text) <> "" Then
            '    tRucC.Text = tRucF.Text
            '    tCiC.Text = ""
            '    BuscoClienteCIRUC Trim(tRucC.Text), lClienteC, RUC:=True
            'End If
            
            If Val(lClienteC.Tag) <> 0 Then Foco tComentario
    End Select
         
        
End Sub

Private Sub tCiC_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Trim(tCiC.Text) = "" Then tRucC.SetFocus: Exit Sub
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
        AccionBuscarClientes tCiF, lClienteF, TipoCliente.Cliente
        If Val(lClienteF.Tag) <> 0 Then Foco tSerie
    End If
End Sub

Private Sub tCiF_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tCiF.Text) = "" Then tRucF.SetFocus: Exit Sub
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
            MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "Cédula No Válida"
            Screen.MousePointer = 0: Exit Function
        End If
        If Not clsGeneral.CedulaValida(CiRuc) Then
            MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "Cédula No Válida"
            Screen.MousePointer = 0: Exit Function
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
    clsGeneral.OcurrioError "Error al buscar el cliente." & Err.Description
End Function

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tUsuario
End Sub


Private Sub tDocumento_Change()
    lDoc.Caption = "": lDoc.Tag = 0
End Sub

Private Sub tDocumento_GotFocus()
    tDocumento.SelStart = 0: tDocumento.SelLength = (Len(tDocumento.Text))
End Sub

Private Sub tDocumento_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Val(lDoc.Tag) <> 0 Or Trim(tDocumento.Text) = "" Then
            tCiC.SetFocus: Exit Sub
        End If
        On Error GoTo errDoc
        
        Dim adQ As Integer, adCodigo As Long, adTexto As String, adTextoFMT As String
        
        Screen.MousePointer = 11
        Dim mDSerie As String, mDNumero As String
        mDNumero = tDocumento.Text
        
        If InStr(tDocumento.Text, "-") <> 0 Then
            mDSerie = Mid(tDocumento.Text, 1, InStr(tDocumento.Text, "-") - 1)
            mDNumero = Mid(tDocumento.Text, InStr(tDocumento.Text, "-") + 1)
        ElseIf Not IsNumeric(Left(tDocumento.Text, 1)) Then
            mDSerie = Left(tDocumento.Text, 1)
            mDNumero = Mid(tDocumento.Text, 2)
        End If
        
        adQ = 0
        Cons = "Select * from Documento " & _
                   " Where DocNumero = " & Val(mDNumero) & _
                   " And DocTipo In (" & TipoDocumento.Credito & ", " & TipoDocumento.Contado & ", " & TipoDocumento.ReciboDePago & ")"
        If Trim(mDSerie) <> "" Then Cons = Cons & " And DocSerie = '" & mDSerie & "'"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            adCodigo = RsAux!DocCodigo
            adTexto = zDocumento(RsAux!DocTipo, RsAux!DocSerie, RsAux!DocNumero, adTextoFMT)
            adQ = 1
            RsAux.MoveNext: If Not RsAux.EOF Then adQ = 2
        End If
        RsAux.Close
        
        Select Case adQ
            Case 2
                Dim miLDocs As New clsListadeAyuda
                Cons = "Select DocCodigo, DocFecha as Fecha, DocSerie + Convert(char(7),DocNumero) as Numero " & _
                           " from Documento " & _
                           " Where DocNumero = " & Val(mDNumero) & _
                           " And DocTipo In (" & TipoDocumento.Credito & ", " & TipoDocumento.Contado & ", " & TipoDocumento.ReciboDePago & ")"
                If Trim(mDSerie) <> "" Then Cons = Cons & " And DocSerie = '" & mDSerie & "'"
                adCodigo = miLDocs.ActivarAyuda(cBase, Cons, 4100, 1)
                Me.Refresh
                If adCodigo <> 0 Then adCodigo = miLDocs.RetornoDatoSeleccionado(0)
                Set miLDocs = Nothing
                
                If adCodigo > 0 Then
                    Cons = "Select * from Documento Where DocCodigo = " & adCodigo
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If Not RsAux.EOF Then
                        adTexto = zDocumento(RsAux!DocTipo, RsAux!DocSerie, RsAux!DocNumero, adTextoFMT)
                    End If
                    RsAux.Close
                End If
        End Select
        
        If adCodigo > 0 Then
            tDocumento.Text = adTextoFMT
            lDoc.Tag = adCodigo: lDoc.Caption = adTexto
        Else
            lDoc.Caption = " No Existe !!"
        End If
        
        If Val(lDoc.Tag) <> 0 Then tCiC.SetFocus
        
        Screen.MousePointer = 0
    End If
    
    Exit Sub
errDoc:
    clsGeneral.OcurrioError "Error al buscar el documento.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function zDocumento(Tipo As Integer, Serie As String, Numero As Long, retSerieNumero As String) As String

    Select Case Tipo
        Case 1: zDocumento = "Ctdo. "
        Case 2: zDocumento = "Créd. "
        Case 3: zDocumento = "N/Dev. "
        Case 4: zDocumento = "N/Créd. "
        Case 5: zDocumento = "Recibo "
        Case 10: zDocumento = "N/Esp. "
    End Select
    
    zDocumento = zDocumento & Trim(Serie) & "-" & Numero
    retSerieNumero = Trim(Serie) & "-" & Numero

End Function
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
                Foco tDocumento
            End If
        Else
            Foco tDocumento
        End If
    End If

End Sub

Private Sub tFPago_Change()
    tFPago.Tag = ""
End Sub

Private Sub tFPago_GotFocus()
    tFPago.Text = tFPago.Tag
    tFPago.SelStart = 0: tFPago.SelLength = Len(tFPago.Text)
End Sub

Private Sub tFPago_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tFPago.Tag) = "" Then
             If IsDate(tFPago.Text) Then
                If CDate(tFPago.Text) < Date Then
                    If Abs(DateDiff("m", CDate(tFPago.Text), Date)) > 3 Then
                        tFPago.Text = Format(DateAdd("yyyy", 1, CDate(tFPago.Text)), "dd/mm/yyyy")
                    End If
                End If
                
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
    On Error Resume Next

    Select Case KeyCode
        Case vbKeyF4
            tCiC.Text = ""
            AccionBuscarClientes tRucC, lClienteC, TipoCliente.Empresa
            If Val(lClienteC.Tag) <> 0 Then Foco tComentario
            
        Case vbKeyMultiply
            If Val(lClienteF.Tag) = 0 Then Exit Sub
            CargoDatosCliente Val(lClienteF.Tag), 2
            
            'If Trim(tCiF.Text) <> "" Then
            '    tCiC.Text = tCiF.Text
            '    tRucC.Text = ""
            '    BuscoClienteCIRUC Trim(tCiC.Text), lClienteC, CI:=True
            'End If
            
            'If Trim(tRucF.Text) <> "" Then
            '    tRucC.Text = tRucF.Text
            '    tCiC.Text = ""
            '    BuscoClienteCIRUC Trim(tRucC.Text), lClienteC, RUC:=True
            'End If
            
            If Val(lClienteC.Tag) <> 0 Then Foco tComentario
    End Select

    
End Sub

Private Sub tRucC_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tRucC.Text) = "" Then tCiC.SetFocus: Exit Sub
        tCiC.Text = ""
        
        Screen.MousePointer = 11
        If BuscoClienteCIRUC(Trim(tRucC.Text), lClienteC, RUC:=True) Then Foco tComentario
        
        Screen.MousePointer = 0
    End If

End Sub

Private Sub tRucF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        On Error Resume Next
        tCiF.Text = ""
        AccionBuscarClientes tRucF, lClienteF, TipoCliente.Empresa
        If Val(lClienteF.Tag) <> 0 Then Foco tSerie
    End If
End Sub

Private Sub tRucF_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tRucF.Text) = "" Then tCiF.SetFocus: Exit Sub
        tCiF.Text = ""
        
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

Private Function FormateoFecha(Fecha As String) As String
    FormateoFecha = Format(Fecha, "d") & " de " & Format(Fecha, "Mmmm") & " de " & Format(Fecha, "yyyy")
End Function

Private Sub AccionGrabar()

    If Not ValidoDatos Then Exit Sub
    
    On Error GoTo errorBT
    Screen.MousePointer = 11
    
    Dim mMoneda As Long
    Dim mIDDisponibilidad As Long
    
    mMoneda = cMoneda.ItemData(cMoneda.ListIndex)
    mIDDisponibilidad = dis_DisponibilidadPara(paCodigoDeSucursal, mMoneda)
    If mIDDisponibilidad = 0 Then
        MsgBox "Su sucursal no puede recibir un cheque en " & Trim(cMoneda.Text) & "." & vbCrLf & _
                    "No hay una disponibilidad para realizar los movimientos de caja que correspondan." & vbCrLf & vbCrLf & _
                    "Consulte con el administrador del sistema.", vbExclamation, "Falta Disponibilidad para la Moneda del Cheque"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    'Valido que no exista un cheque con las mismas caracteristicas---------------------------------------------------------
    'Hay que validar que el cheque no este ingresado (Banco, Sucursal, Numero)
    Cons = "Select * from ChequeDiferido " _
           & " Where CDiBanco = " & ZBancoSucursal(tBanco.Tag, Banco:=True) _
           & " And CDiSucursal = " & ZBancoSucursal(tBanco.Tag, Sucursal:=True) _
           & " And CDiSerie = '" & Trim(tSerie.Text) & "'" _
           & " And CDiNumero = " & tNumero.Text & " AND CDiEliminado IS NULL"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        Screen.MousePointer = 0
        MsgBox "El cheque ya ha sido ingresado al sistema." & Chr(vbKeyReturn) _
                & "Verifique los datos ingresados en Seguimiento de Cheques.", vbCritical, "Cheque Ingresado !!!"
        RsAux.Close
        Exit Sub
    End If
    RsAux.Close
    '--------------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 0
    
    'Valido si se ingreso un documento
    Dim mButtons As Long
    Dim bSuceso As Boolean: bSuceso = False
    
    Dim prmSucesoUsr As Long, prmSucesoDef As String
    
    aTexto = "Confirma grabar los datos del cheque."
    mButtons = vbQuestion + vbYesNo
    
    If Val(lDoc.Tag) = 0 Then
        If Not tFPago.Visible Then
            aTexto = "Falta ingresar el documento asociado al cheque." & vbCrLf & _
                         "Graba igualmente ?"
        Else
            aTexto = "Falta ingresar el documento asociado al cheque." & vbCrLf & _
                         "En los cheques diferidos se debe ingresar el documento." & vbCrLf & vbCrLf & _
                         "Graba igualmente ?"
            mButtons = vbQuestion + vbYesNo + vbDefaultButton2
            bSuceso = True
        End If
    End If
    
    If tFPago.Visible Then
        '04/03/2020 valido que el vencimiento no sea mayor a 180 días.
        Dim dVence As Date
        dVence = CDate(tFPago.Tag)
        Dim topeDias As Byte
        topeDias = IIf(Weekday(dVence) = vbSunday, 178, IIf(Weekday(dVence) = vbSaturday, 179, 180))
        
        If (Abs(DateDiff("d", CDate(tFEmision.Tag), dVence)) > topeDias) Then
            MsgBox "ATENCIÓN!!!, la fecha de vencimiento del cheque no puede ser mayor a " & DateAdd("d", topeDias, CDate(tFEmision.Tag)) & vbCrLf & vbCrLf & "Imposible grabar", vbCritical, "ATENCIÓN"
            Exit Sub
        End If
        
    End If
    
    If MsgBox(aTexto, mButtons, "Grabar Cheque") = vbNo Then
        If Val(lDoc.Tag) = 0 Then Foco tDocumento
        Exit Sub
    End If
    
    If bSuceso Then
        'Llamo al registro del Suceso-------------------------------------------------------------
        Dim objSuceso As New clsSuceso
        objSuceso.ActivoFormulario paCodigoDeUsuario, "Cheque Diferido s/Documento", cBase
        Me.Refresh
        prmSucesoUsr = objSuceso.RetornoValor(Usuario:=True)
        prmSucesoDef = objSuceso.RetornoValor(Defensa:=True)
        Set objSuceso = Nothing
        If prmSucesoUsr = 0 Then Screen.MousePointer = 0: Exit Sub
        '---------------------------------------------------------------------------------------------
    End If
    
    Screen.MousePointer = 11
    FechaDelServidor
    
    
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET


    If Not bYaInserteChequeLC And Me.prmTAG_LiqCamion > 0 Then
        'consulto para validar que no tenga utlizado ya el tag.
        Cons = "SELECT CDiCodigo FROM ChequeDiferido WHERE CDiTag = " & Me.prmTAG_LiqCamion
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            On Error Resume Next
            RsAux.Close
            cBase.RollbackTrans
            Screen.MousePointer = 0
            MsgBox "El TAG identificatorio del cheque para la liquidación de camioneros ya fue utilizado, debera cancelar la ejecución y reingresar al formulario.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
        RsAux.Close
    End If



    Cons = "Select * from ChequeDiferido Where CDiCodigo = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    RsAux.AddNew
    RsAux!CDiBanco = ZBancoSucursal(tBanco.Tag, Banco:=True)
    RsAux!CDiSucursal = ZBancoSucursal(tBanco.Tag, Sucursal:=True)
    RsAux!CDiSerie = Trim(UCase(tSerie.Text))
    RsAux!CDiNumero = tNumero.Text
    RsAux!CDiMoneda = cMoneda.ItemData(cMoneda.ListIndex)
    RsAux!CDiImporte = CCur(tImporte.Text)
    
    If tFPago.Visible Then RsAux!CDiVencimiento = Format(tFPago.Tag, "mm/dd/yyyy") Else RsAux!CDiVencimiento = Null
    
    RsAux!CDiLibrado = Format(tFEmision.Tag, "mm/dd/yyyy")
    RsAux!CDiCliente = lClienteC.Tag
    RsAux!CDiUsuario = tUsuario.Tag
    RsAux!CDiIngresado = Format(gFechaServidor, "mm/dd/yyyy hh:mm")
    RsAux!CDiClienteFactura = lClienteF.Tag
    If Trim(tComentario.Text) <> "" Then RsAux!CDiComentario = Trim(tComentario.Text)
    
    If Val(lDoc.Tag) <> 0 Then RsAux!CDiDocumento = Val(lDoc.Tag)
    
    'Si tengo prendido el tag de liquidación de camioneros entonces grabo el mismo.
    If prmTAG_LiqCamion > 0 Then RsAux("CDiTag") = prmTAG_LiqCamion
    
    RsAux.Update
    RsAux.Close
    
    'Si el Cheque es Diferido ---> Hago mov Caja
    If tFPago.Visible Then
        'Hago Movimiento de Caja  (Salida -- xq' ingreso un cheque).
        aTexto = "Cheque " & tSerie.Text & "-" & tNumero.Text & " (" & tBanco.FormattedText & ") " & Trim(lBanco.Caption)
        MovimientoDeCaja paMCChequeDiferido, gFechaServidor, mIDDisponibilidad, CInt(mMoneda), CCur(tImporte.Text), aTexto, True
    End If
    
    If bSuceso Then
        'Registro el Suceso -----------------------------------------------------------------------------------------------------------
        aTexto = tBanco.FormattedText & " " & Trim(tSerie.Text) & "-" & Trim(tNumero.Text)
        clsGeneral.RegistroSuceso cBase, gFechaServidor, prmSucesoCheque, paCodigoDeTerminal, prmSucesoUsr, 0, _
                                Descripcion:="Ch.Dif. s/Doc. " & aTexto, _
                                Defensa:=Trim(prmSucesoDef), Valor:=CCur(tImporte.Text), idCliente:=Val(lClienteC.Tag)

    End If
    
    cBase.CommitTrans    'Fin de la TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    bYaInserteChequeLC = True
    Screen.MousePointer = 0
    
    If prmTipo <> -1 Then
        If MsgBox("Desea ingresar un nuevo cheque.", vbQuestion + vbYesNo + vbDefaultButton2, "Ingresa Otro Cheque ?") = vbYes Then
            If tFPago.Visible Then
                LimpioCampos bNew:=True
                Foco tNumero
            Else
                LimpioCampos
                tCiF.Text = "": tCiC.Text = ""
                tRucF.Text = "": tRucC.Text = ""
                lClienteF.Caption = "": lClienteC.Caption = ""
                lClienteF.Tag = 0: lClienteC.Tag = 0
                tCiF.SetFocus
            End If
            lIdCheque.Caption = "CHEQUE Nº " & Val(lIdCheque.Tag) + 1
            lIdCheque.Tag = Val(lIdCheque.Tag) + 1
            
        Else
            sNuevo = False: Unload Me
        End If
    Else
        LimpioCampos
        prmTipo = -1
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
        MsgBox "Ingrese su dígito de usuario.", vbExclamation, "Faltan Datos"
        Foco tUsuario: Exit Function
    End If
    
    If Val(lClienteF.Tag) = 0 Then
        MsgBox "Se debe ingresar el cliente titular de la operación.", vbExclamation, "Faltan Datos"
        tCiF.SetFocus: Exit Function
    End If
    
    If Val(lClienteC.Tag) = 0 Then
        MsgBox "Se debe ingresar el cliente que emite el cheque.", vbExclamation, "Faltan Datos"
        tCiC.SetFocus: Exit Function
    End If
    
    If Trim(tSerie.Text) = "" Then
        MsgBox "La serie del cheque no es correcta. Verifique", vbExclamation, "Faltan Datos"
        Foco tSerie: Exit Function
    End If
    
    If Not IsNumeric(tNumero.Text) Then
        MsgBox "El número del cheque no es correcto. Verifique", vbExclamation, "Faltan Datos"
        Foco tNumero: Exit Function
    End If
    
    If Not IsNumeric(tImporte.Text) Then
        MsgBox "El importe del cheque no es correcto. Verifique", vbExclamation, "Faltan Datos"
        Foco tImporte: Exit Function
    End If
    
    If tBanco.Tag = "" Then
        MsgBox "El banco seleccionado no es correcto. Verifique", vbExclamation, "Faltan Datos"
        Foco tBanco: Exit Function
    End If
    
    If tFPago.Visible Then
        If Not IsDate(tFPago.Tag) Then
            MsgBox "La fecha de vencimiento ingresada no es correcta. Verifique", vbExclamation, "Faltan Datos"
            Foco tFPago: Exit Function
        End If
    End If
    
    If Not IsDate(tFEmision.Tag) Then
        MsgBox "La fecha de emisión del cheque no es correcta. Verifique", vbExclamation, "Faltan Datos"
        Foco tFEmision: Exit Function
    End If
    
    If tFPago.Visible Then
        If CDate(tFPago.Tag) < CDate(tFEmision.Tag) Then
            MsgBox "La fecha de emisión del cheque no debe ser mayor a la del pago. Verifique", vbExclamation, "ATENCIÓN"
            Foco tFEmision: Exit Function
        End If
    
        FechaDelServidor
        'Valido que sea diferido
        If Format(tFPago.Tag, "dd/mm/yyyy") = Format(gFechaServidor, "dd/mm/yyyy") Then
            MsgBox "El cheque ingresado NO ES DIFERIDO, la fecha del vencimiento es la del día de hoy, por lo tanto ....  es un CHEQUE AL DÍA.", vbExclamation, "ATENCIÓN"
            Foco tFPago: Exit Function
        End If
    
        If Format(tFPago.Tag, "yyyy/mm/dd") < Format(gFechaServidor, "yyyy/mm/dd") Then
            MsgBox "La fecha de pago del cheque es menor al día de hoy. Verifique", vbExclamation, "ATENCIÓN"
            Foco tFPago: Exit Function
        End If
    End If
    
    'Si es un cheque al Dia , la fecha de emision debe ser la del dia de HOY IRMA: 22/11/2002
    If Not tFPago.Visible Then
        If Format(CDate(tFEmision.Tag), "dd/mm/yyyy") <> Format(gFechaServidor, "dd/mm/yyyy") Then
            MsgBox "La fecha de emisión del cheque debe ser igual a hoy" & vbCrLf & "No se pueden ingresar cheques al día que no estén librados 'hoy'.", vbExclamation, "Fecha de Librado"
            Foco tFPago: Exit Function
        End If
    End If
    
    ValidoDatos = True
    Exit Function

errValido:
    clsGeneral.OcurrioError "Error al validar los datos. " & Err.Description
    Screen.MousePointer = 0
End Function

Private Sub CargoDatosCliente(Codigo As Long, Optional Cual As Integer = 1)

    On Error Resume Next
    Dim mNombre As String, mCi As String, mRuc As String
    mNombre = "": mCi = "": mRuc = ""
    
    If Cual = 1 Then
        lClienteF.Caption = ""
        tRucF.Text = "": tCiF.Text = ""
    Else
        lClienteC.Caption = ""
        tRucC.Text = "": tCiC.Text = ""
    End If
    
    Cons = "Select * from Cliente, CEmpresa Where CliCodigo = " & Codigo & " And CliCodigo = CEmCliente"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!CEmFantasia) Then mNombre = Trim(RsAux!CEmFantasia)
        If Not IsNull(RsAux!CEmNombre) Then mNombre = mNombre & " (" & Trim(RsAux!CEmNombre) & ")"
        If Not IsNull(RsAux!CliCIRuc) Then mRuc = Trim(RsAux!CliCIRuc)
    End If
    RsAux.Close
    
    If Trim(mNombre) = "" Then
        Cons = "Select * from Cliente, CPersona Where CliCodigo = " & Codigo & " And CliCodigo = CPeCliente"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux.EOF Then
            mNombre = ArmoNombre(RsAux!CPeApellido1, Format(RsAux!CPeApellido2, "#"), RsAux!CPeNombre1, Format(RsAux!CPeNombre2, "#"))
            If Not IsNull(RsAux!CliCIRuc) Then mCi = clsGeneral.RetornoFormatoCedula(RsAux!CliCIRuc)
        End If
        RsAux.Close
    End If
    
    If Cual = 1 Then
        If Trim(mCi) <> "" Then tCiF.Text = mCi
        If Trim(mRuc) <> "" Then tRucF.Text = mRuc
        lClienteF.Caption = " " & mNombre
        lClienteF.Tag = Codigo
    Else
        If Trim(mCi) <> "" Then tCiC.Text = mCi
        If Trim(mRuc) <> "" Then tRucC.Text = mRuc
        lClienteC.Caption = " " & mNombre
        lClienteC.Tag = Codigo
    End If
    
End Sub

Private Sub AccionBuscarClientes(txtCiRuc As Control, lblNombre As Control, aTipoCliente As Integer)
    
    Screen.MousePointer = 11
    On Error GoTo errCargar
    Dim aIdCliente As Long, aTipo As Integer
    Dim objBuscar As New clsBuscarCliente
    
    If aTipoCliente = TipoCliente.Cliente Then objBuscar.ActivoFormularioBuscarClientes cBase, Persona:=True
    If aTipoCliente = TipoCliente.Empresa Then objBuscar.ActivoFormularioBuscarClientes cBase, Empresa:=True
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
        
        If aTipo = TipoCliente.Cliente Then
            Cons = "Select * from Cliente, CPersona Where CliCodigo = " & aCodigo & " And CliCodigo = CPeCliente"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux.EOF Then
                lblNombre.Caption = " " & ArmoNombre(RsAux!CPeApellido1, Format(RsAux!CPeApellido2, "#"), RsAux!CPeNombre1, Format(RsAux!CPeNombre2, "#"))
                If Not IsNull(RsAux!CliCIRuc) Then txtCiRuc.Text = clsGeneral.RetornoFormatoCedula(RsAux!CliCIRuc) Else txtCiRuc.Text = ""
                lblNombre.Tag = aCodigo
            End If
            RsAux.Close
        End If
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar los datos del cliente.", Err.Description
End Sub


Private Function ZBancoSucursal(Texto As String, Optional Banco As Boolean = False, Optional Sucursal As Boolean = False)
    If Banco Then ZBancoSucursal = Mid(Texto, 1, InStr(Texto, "-") - 1)
    If Sucursal Then ZBancoSucursal = Mid(Texto, InStr(Texto, "-") + 1, Len(Texto))
End Function


Private Sub EstadoCampos(Estado As Boolean)

    cTipo.Enabled = Not Estado
    
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
    tDocumento.Enabled = Estado
    
    Dim bkColor  As Long
    
    If Estado Then
        bkColor = RGB(250, 235, 215)
        bkColor = RGB(255, 228, 196)
        bkColor = RGB(255, 235, 205)
        'bkColor = RGB(255, 239, 213)
        
        bkColor = RGB(255, 228, 196)
        bkColor = RGB(255, 218, 185)
    Else
        bkColor = Me.BackColor
        bkColor = RGB(255, 228, 196)
    End If
    
    lClienteF.BackColor = bkColor
    lClienteC.BackColor = bkColor
    tComentario.BackColor = bkColor
    tUsuario.BackColor = bkColor
    
    tCiF.BackColor = bkColor
    tRucF.BackColor = bkColor
    tCiC.BackColor = bkColor
    tRucC.BackColor = bkColor

End Sub

Private Sub LimpioCampos(Optional bNew As Boolean = False)

On Error Resume Next
    
    If bNew Then
        If IsNumeric(tNumero.Text) Then tNumero.Text = Val(tNumero.Text) + 1
        
        aTexto = Format(DateAdd("m", 1, CDate(tFPago.Tag)), "dd/mm/yyyy")
        tFPago.Text = FormateoFecha(aTexto)
        tFPago.Tag = Format(aTexto, "dd/mm/yyyy")
        Exit Sub
    End If
    
    tImporte.Text = ""
    lImporte.Caption = Space(400)
    tFPago.Text = ""
    tNumero.Text = ""
    tSerie.Text = ""
    tBanco.Text = "": lBanco.Caption = ""
    tFEmision.Text = ""
    tDocumento.Text = "": lDoc.Caption = ""
    tComentario.Text = ""
    
End Sub

Private Sub tSerie_LostFocus()
On Error Resume Next
    tSerie.Text = UCase(tSerie.Text)
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        On Error Resume Next
        If IsNumeric(tUsuario.Text) Then tUsuario.Tag = z_BuscoUsuarioDigito(Val(tUsuario.Text), Codigo:=True)
        If Val(tUsuario.Tag) <> 0 Then AccionGrabar
    End If
    
End Sub


Function ArmoNombre(Ape1 As String, Ape2 As String, Nom1 As String, Nom2 As String) As String

    ArmoNombre = Trim(Ape1) & " " & Trim(Ape2)
    ArmoNombre = Trim(ArmoNombre) & ", " & Trim(Nom1) & " " & Trim(Nom2)
    
End Function
