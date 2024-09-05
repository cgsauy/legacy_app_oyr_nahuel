VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMaCheque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seguimiento de Cheques"
   ClientHeight    =   4575
   ClientLeft      =   2880
   ClientTop       =   3210
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMaCheque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7935
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Canjear cheque"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   550
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "rebotar"
            Object.ToolTipText     =   "Rebotar cheque."
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "desdepositar"
            Object.ToolTipText     =   "Desdepositar"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   4150
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Clientes"
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   120
      TabIndex        =   27
      Top             =   3480
      Width           =   7695
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   795
      End
      Begin VB.Label lClienteC 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Rodriguez Fernandez, Rodrigo Bernardino"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1080
         TabIndex        =   30
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   6495
      End
      Begin VB.Label lClienteD 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "21 378350 0011"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1080
         TabIndex        =   29
         Top             =   255
         UseMnemonic     =   0   'False
         Width           =   6495
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Documento:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   255
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Generales"
      ForeColor       =   &H00800000&
      Height          =   1575
      Left            =   120
      TabIndex        =   25
      Top             =   540
      Width           =   7695
      Begin VB.TextBox tDocumento 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   840
         MaxLength       =   12
         TabIndex        =   13
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox tComentario 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         MaxLength       =   50
         TabIndex        =   11
         Top             =   900
         Width           =   6735
      End
      Begin VB.TextBox tCSerie 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1620
         MaxLength       =   2
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox tCNumero 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2040
         MaxLength       =   7
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox tCVencimiento 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   840
         MaxLength       =   12
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox tCLibrado 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2820
         MaxLength       =   12
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox tCImporte 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4980
         MaxLength       =   12
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin MSMask.MaskEdBox tBanco 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   8388608
         PromptInclude   =   0   'False
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99-999"
         PromptChar      =   "_"
      End
      Begin VB.Label lDoc 
         BackColor       =   &H00808080&
         Caption         =   "Label9"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1980
         TabIndex        =   32
         Top             =   1200
         Width           =   1995
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "&Doc.:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "&Observ.:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lBanco 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S/D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2820
         TabIndex        =   26
         Top             =   240
         Width           =   4755
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "&Número:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   280
         Width           =   615
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "&Vence:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   615
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "&Librado:"
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   660
         Width           =   735
      End
      Begin VB.Label lImporte 
         BackStyle       =   0  'Transparent
         Caption         =   "&Importe:"
         Height          =   255
         Left            =   4260
         TabIndex        =   8
         Top             =   660
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información del Cheque"
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   120
      TabIndex        =   15
      Top             =   2220
      Width           =   7695
      Begin VB.Label lRebotado 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingresado:"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3480
         TabIndex        =   36
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label lRebotadoT 
         BackStyle       =   0  'Transparent
         Caption         =   "Rebotado:"
         Height          =   255
         Left            =   2700
         TabIndex        =   35
         Top             =   570
         Width           =   855
      End
      Begin VB.Label lEliminado 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingresado:"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5880
         TabIndex        =   34
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label lEliminadoT 
         BackStyle       =   0  'Transparent
         Caption         =   "Eliminado:"
         Height          =   255
         Left            =   5100
         TabIndex        =   33
         Top             =   570
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   5220
         TabIndex        =   24
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Depositado:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Ingresado:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   570
         Width           =   855
      End
      Begin VB.Label lMoneda 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "N/D"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1080
         TabIndex        =   20
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lUsuario 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Usuario:"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5880
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lFDepositado 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Depositado:"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1080
         TabIndex        =   18
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lFIngresado 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Ingresado:"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1080
         TabIndex        =   17
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label lBancoDeposito 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Banco:"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2640
         TabIndex        =   16
         Top             =   840
         Width           =   4935
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7440
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCheque.frx":0442
            Key             =   "eliminar"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCheque.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCheque.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCheque.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCheque.frx":0A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCheque.frx":0DAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCheque.frx":10C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCheque.frx":13E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCheque.frx":15BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMaCheque.frx":1794
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuNuevo 
         Caption         =   "Ca&njear Cheque"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
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
      Begin VB.Menu MnuExit 
         Caption         =   "Del formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmMaCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type typData
    Moneda As Integer
    Importe As Currency
    NumeroCheque As String
    Documento As Long
    Vencimiento As Date
    Librado As Date
    IdSucursalDepositado As Long
    bDepositado As Boolean
    bRebotado As Boolean
End Type

Dim myData As typData

Public prmIdCheque As Long
Dim prmIdCliente As Long
Dim sModificar As Boolean
Dim gSucesoUsr As Long, gSucesoDef As String

Dim mTexto As String
Dim rsX As rdoResultset

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()

    On Error GoTo ErrLoad
    FechaDelServidor
    
    LimpioFicha True
    EstadoIngreso False
    EstadoBotones
    
    If prmIdCheque <> 0 Then CargoCheque prmIdCheque
    
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Error al inicialiar el formulario.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    EndMain
End Sub

Private Sub Label15_Click()
    tBanco.SelStart = 0
    tBanco.SelLength = 5
    tBanco.SetFocus
End Sub

Private Sub Label17_Click()
    Foco tCLibrado
End Sub

Private Function zNombreCliente(Codigo As Long) As String

Dim aStr As String

    On Error GoTo errCliente
    Cons = "Select Tipo = 1, CliCiRuc, CliTipo, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2) " _
           & " From Cliente, CPersona " _
           & " Where CliCodigo = " & Codigo _
           & " And CliCodigo = CPeCliente " _
                                                & " UNION " _
           & " Select Tipo = 2, CliCiRuc, CliTipo, Nombre = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')')  " _
           & " From Cliente, CEmpresa " _
           & " Where CliCodigo = " & Codigo _
           & " And CliCodigo = CEmCliente"

    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        If RsAux!Tipo = 1 Then
            If Not IsNull(RsAux!CliCIRuc) Then aStr = " (" & clsGeneral.RetornoFormatoCedula(RsAux!CliCIRuc) & ") "
        Else
            If Not IsNull(RsAux!CliCIRuc) Then aStr = " (" & clsGeneral.RetornoFormatoRuc(RsAux!CliCIRuc) & ") "
        End If
    End If
    zNombreCliente = aStr & Trim(RsAux!Nombre)
    
    RsAux.Close
    Exit Function
    
errCliente:
    zNombreCliente = ""
    clsGeneral.OcurrioError "Error al cargar los datos del cliente.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuEliminar_Click()
    AccionEliminar
End Sub

Private Sub MnuExit_Click()
    Unload Me
End Sub

Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuModificar_Click()
    AccionModificar
End Sub

Private Sub MnuNuevo_Click()
    AccionNuevo
End Sub

Private Sub tBanco_Change()
    tBanco.Tag = ""
End Sub

Private Sub tBanco_GotFocus()
    tBanco.SelStart = 0: tBanco.SelLength = 6
End Sub

Private Sub tBanco_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo errKD
    
    Select Case KeyCode
        Case vbKeyF1
        
            Screen.MousePointer = 11
            Cons = "Select BanCodigoB as 'Nº B', SBaCodigoS as 'Nº S', BanNombre as 'Banco', SBaNombre as 'Sucursal' " & _
                        " From BancoSSFF, SucursalDeBanco" & _
                        " Where SBaBanco = BanCodigo"
                        
            If Len(tBanco.Text) = 2 Then Cons = Cons & " And BanCodigoB = " & tBanco.Text
            Cons = Cons & " Order by BanNombre, SBaNombre"
                
            Dim objHelp As New clsListadeAyuda, mSel As String
            mSel = ""
            If objHelp.ActivarAyuda(cBase, Cons, Titulo:="Lista de Bancos") Then
                mSel = Format(objHelp.RetornoDatoSeleccionado(0), "00")
                mSel = mSel & Format(objHelp.RetornoDatoSeleccionado(1), "000")
            End If
            Set objHelp = Nothing
            
            Me.Refresh
            Screen.MousePointer = 0
            If mSel <> "" Then
                tBanco.Text = mSel
                Call tBanco_KeyPress(vbKeyReturn)
            End If

    End Select
    
    Exit Sub
errKD:
    clsGeneral.OcurrioError "Error al procesar la lista de bancos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tBanco_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Len(tBanco.Text) = 5 Then
            If BuscoBancoEmisor(tBanco.Text) Then Foco tCSerie
        Else
            If Not sModificar Then Foco tCSerie
        End If
    End If
        
End Sub

Private Sub tCImporte_GotFocus()
    tCImporte.SelStart = 0
    tCImporte.SelLength = Len(tCImporte.Text)
End Sub

Private Sub tCImporte_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tCImporte.Text) Then
            tCImporte.Text = Format(tCImporte.Text, FormatoMonedaP)
            Foco tComentario
        End If
    End If
    
End Sub

Private Sub tCLibrado_GotFocus()
    tCLibrado.SelStart = 0
    tCLibrado.SelLength = Len(tCLibrado.Text)
End Sub

Private Sub tCLibrado_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If IsDate(tCLibrado.Text) Then
            tCLibrado.Text = Format(tCLibrado.Text, "d-Mmm yyyy")
            Foco tCImporte
        End If
    End If
    
End Sub

Private Sub tCNumero_GotFocus()
    tCNumero.SelStart = 0
    tCNumero.SelLength = Len(tCNumero.Text)
End Sub

Private Sub tCNumero_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tCNumero.Text) And Not sModificar Then BuscoCheque
        If Not sModificar Then Foco tBanco: Exit Sub
        If tCVencimiento.Enabled Then Foco tCVencimiento Else Foco tCLibrado
    End If
    
End Sub

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tDocumento.Enabled Then Foco tDocumento Else AccionGrabar
    End If
End Sub

Private Sub tCSerie_GotFocus()
    tCSerie.SelStart = 0: tCSerie.SelLength = Len(tCSerie.Text)
End Sub

Private Sub tCSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then Foco tCNumero
End Sub

Private Sub tCVencimiento_GotFocus()
    tCVencimiento.SelStart = 0
    tCVencimiento.SelLength = Len(tCVencimiento.Text)
End Sub

Private Sub tCVencimiento_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If IsDate(tCVencimiento.Text) Then
            tCVencimiento.Text = Format(tCVencimiento.Text, "d-Mmm yyyy")
            Foco tCLibrado
        End If
        If Trim(tCVencimiento.Text) = "" And myData.Vencimiento = CDate("1/1/1900") Then Foco tCLibrado
    End If
    
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
            AccionGrabar
            Exit Sub
        End If
        On Error GoTo errDoc
'        If Not IsNumeric(tDocumento.Text) Then Exit Sub
        
        Dim adQ As Integer, adCodigo As Long, adTexto As String, adTextoFmt As String
        
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
                   " And DocTipo In (" & TipoDocumento.Contado & ", " & TipoDocumento.ReciboDePago & ")"
        If Trim(mDSerie) <> "" Then Cons = Cons & " And DocSerie = '" & mDSerie & "'"
        
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            adCodigo = RsAux!DocCodigo
            adTexto = zDocumento(RsAux!DocTipo, RsAux!DocSerie, RsAux!DocNumero, adTextoFmt)
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
                           " And DocTipo In (" & TipoDocumento.Contado & ", " & TipoDocumento.ReciboDePago & ")"
                If Trim(mDSerie) <> "" Then Cons = Cons & " And DocSerie = '" & mDSerie & "'"
                adCodigo = miLDocs.ActivarAyuda(cBase, Cons, 4100, 1)
                Me.Refresh
                If adCodigo <> 0 Then adCodigo = miLDocs.RetornoDatoSeleccionado(0)
                Set miLDocs = Nothing
                
                If adCodigo > 0 Then
                    Cons = "Select * from Documento Where DocCodigo = " & adCodigo
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If Not RsAux.EOF Then
                        adTexto = zDocumento(RsAux!DocTipo, RsAux!DocSerie, RsAux!DocNumero, adTextoFmt)
                    End If
                    RsAux.Close
                End If
        End Select
        
        If adCodigo > 0 Then
            tDocumento.Text = adTextoFmt
            lDoc.Tag = adCodigo: lDoc.Caption = adTexto
        Else
            lDoc.Caption = " No Existe !!"
        End If
        
        If Val(lDoc.Tag) <> 0 Then AccionGrabar
        
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

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
  
  Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        
        Case "rebotar": AccionDesDepositar bRebota:=True
        Case "desdepositar": AccionDesDepositar bRebota:=False
        
        Case "salir": Unload Me
    End Select
End Sub

Private Function BuscoBancoEmisor(Codigo As String) As Boolean
    
Dim Banco As String, Sucursal As String

    On Error GoTo errCargar
    Screen.MousePointer = 11
    Banco = Mid(Codigo, 1, 2)
    Sucursal = Mid(Codigo, 3, 3)
    BuscoBancoEmisor = True
    
    Cons = "Select * from BancoSSFF, SucursalDeBanco" _
          & "  Where BanCodigoB = " & Banco _
          & "  And SBaCodigoS = " & Sucursal _
          & "  And SBaBanco = BanCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        lBanco.Caption = Trim(RsAux!BanNombre) & " - " & Trim(RsAux!SBaNombre)
        tBanco.Tag = RsAux!BanCodigo & "-" & RsAux!SBaCodigo
    Else
        MsgBox "No existe un banco para el código ingresado.", vbExclamation, "ATENCIÓN"
        tBanco.Tag = ""
        BuscoBancoEmisor = False
    End If
    RsAux.Close
    
    Screen.MousePointer = 0
    Exit Function

errCargar:
    clsGeneral.OcurrioError "Error al cargar el banco emisor.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub AccionNuevo()
    On Error GoTo errorBT
    
    Dim mIDDisponibilidad As Long, aIDMoneda As Long
    
    aIDMoneda = myData.Moneda
    mIDDisponibilidad = dis_DisponibilidadPara(paCodigoDeSucursal, aIDMoneda)
    If mIDDisponibilidad = 0 Then
        MsgBox "Su sucursal no puede recibir un cheque en " & Trim(lMoneda.Caption) & "." & vbCrLf & _
                    "No hay una disponibilidad para realizar los movimientos de caja que correspondan." & vbCrLf & vbCrLf & _
                    "Consulte con el administrador del sistema.", vbExclamation, "Falta Disponibilidad para la Moneda del Cheque"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Dim bAlDia As Boolean
    bAlDia = (myData.Vencimiento = CDate("1/1/1900"))
    
    Dim mMsg As String
    mMsg = "Confirma canjear el cheque por efectivo ?"
    If Not bAlDia And Not myData.bRebotado Then
        mMsg = vbCrLf & "Como el Cheque es Diferido y No está Rebotado, se va a hacer un ingreso de caja por el efectivo."
    End If
    
    'Canjear Cheque por efectivo
    If MsgBox(mMsg, vbQuestion + vbYesNo, "Canjear Cheque") = vbNo Then Exit Sub
    
    'Llamo al registro del Suceso-------------------------------------------------------------
    Dim objSuceso As New clsSuceso
    objSuceso.ActivoFormulario paCodigoDeUsuario, "Canje de Cheques Diferidos", cBase
    Me.Refresh
    gSucesoUsr = objSuceso.RetornoValor(Usuario:=True)
    gSucesoDef = objSuceso.RetornoValor(Defensa:=True)
    Set objSuceso = Nothing
    If gSucesoUsr = 0 Then Screen.MousePointer = 0: Exit Sub
    '---------------------------------------------------------------------------------------------
        
    Screen.MousePointer = 11
        
    FechaDelServidor
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
    
    'Como es un Canje Hay que eliminar el cheque
    '27/03/2003 Nuevo campo CDiEliminado (fecha) de ahora en más se updatea -----------------------------
    Cons = "Select * from ChequeDiferido Where CDiCodigo = " & prmIdCheque
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.Edit
    RsAux!CDiEliminado = Format(gFechaServidor, "mm/dd/yyyy hh:mm")
    RsAux.Update
    RsAux.Close
    '---------------------------------------------------------------------------------------------------------------------
    
    'Registro el Suceso -------------------------------------------------------
    mTexto = tBanco.FormattedText & " " & Trim(tCSerie.Text) & "-" & Trim(tCNumero.Text)
    gSucesoDef = gSucesoDef & vbCrLf & zDatosCheque
    clsGeneral.RegistroSuceso cBase, gFechaServidor, prmSucesoCheque, paCodigoDeTerminal, gSucesoUsr, Val(lDoc.Tag), _
                            Descripcion:="Canje x Efec. " & mTexto, _
                            Defensa:=Trim(gSucesoDef), Valor:=CCur(tCImporte.Text), idCliente:=Val(lClienteC.Tag)
    
    'Hago el ingreso de Caja
    'Si dif y no está rebotado
    If Not bAlDia And Not myData.bRebotado Then
        MovimientoDeCaja paMCChequeDiferido, gFechaServidor, mIDDisponibilidad, myData.Moneda, CCur(tCImporte.Text), "Canje de Cheques Diferidos." & mTexto
    End If
       
    cBase.CommitTrans    'Fin de la TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    LimpioFicha True
    Botones False, False, False, False, False, Toolbar1, Me
    Dim mID_Aux As Long: mID_Aux = prmIdCheque
    CargoCheque mID_Aux
    
    tBanco.SetFocus
    Screen.MousePointer = 0
    Exit Sub

errorBT:
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionEliminar()

'NO SE PUEDE ELIMINAR UN CHEQUE CUANDO ESTÁ DEPOSITADO (HAY QUE SABER SI REBOTO O NO x MOVS

'   --> Para Eliminar: Si está depositado:   a)  Sacarlo de banco haciendo Gasto o Transferencia a la Inversa
'                                                             b)  Quitar marcas de depósito y marcarlo como rebotado
'                              Si es diferido y no está rebotado (ni depositado) hacer movs de caja para compensar
'                              Marcarlo como eliminado.
' Por lo tanto si el cheque está depositado DEBO SABER si para eliminarlo va a rebotar o a desdepositarse.

On Error GoTo errBEliminar

    Dim mIDDisponibilidad As Long, aIDMoneda As Integer
    
    Dim bAlDia As Boolean
    bAlDia = (myData.Vencimiento = CDate("1/1/1900"))
    
    aIDMoneda = Val(myData.Moneda)   'Moneda
    mIDDisponibilidad = dis_DisponibilidadPara(paCodigoDeSucursal, CLng(aIDMoneda))
    If mIDDisponibilidad = 0 Then
        MsgBox "Su sucursal no puede recibir un cheque en " & Trim(lMoneda.Caption) & "." & vbCrLf & _
                    "No hay una disponibilidad para realizar los movimientos de caja que correspondan." & vbCrLf & vbCrLf & _
                    "Consulte con el administrador del sistema.", vbExclamation, "Falta Disponibilidad para la Moneda del Cheque"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Dim bRespuesta As Integer
    If Not bAlDia And Not myData.bRebotado Then
        mTexto = "Confirma eliminar el cheque del sistema ?." & vbCrLf & _
                        "El cheque es Diferido y no está Rebotado." & vbCrLf & _
                        "Si = Realizar un ingreso de caja para compensar la salida generada y eliminar el cheque." & vbCrLf & _
                        "No = Cancelar la eliminación."
                        
        mTexto = "Confirma eliminar el cheque del sistema ?." & vbCrLf & vbCrLf & _
                        "Presione 'SI' para Ingresar el Cheque a la Caja y Sacarlo del rubro Cheques Diferidos (el sistema va a realizar el movimiento de caja)." & vbCrLf & vbCrLf & _
                        "Presione 'NO' para cancelar." & vbCrLf & vbCrLf & _
                        "Nota: el Cheque es Diferido y No está Rebotado."
        
    Else
        mTexto = "Confirma eliminar el cheque del sistema. " & vbCrLf & vbCrLf & _
                        "NO se van a realizar movimientos de caja y NO SE DEBEN hacer." & vbCrLf & vbCrLf & _
                        "En el único caso que se hace un ingreso de caja es cuando: " & vbCrLf & _
                        "El Cheque Es Diferido y No Está Rebotado"
    End If
    
    bRespuesta = MsgBox(mTexto, vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar Cheque")
    
    If bRespuesta = vbNo Then Exit Sub
    
    'Llamo al registro del Suceso-------------------------------------------------------------
    Dim objSuceso As New clsSuceso
    objSuceso.ActivoFormulario paCodigoDeUsuario, "Eliminación de Cheques Diferidos", cBase
    Me.Refresh
    gSucesoUsr = objSuceso.RetornoValor(Usuario:=True)
    gSucesoDef = objSuceso.RetornoValor(Defensa:=True)
    Set objSuceso = Nothing
    If gSucesoUsr = 0 Then Screen.MousePointer = 0: Exit Sub
    '---------------------------------------------------------------------------------------------
    
    On Error GoTo errorBT
    Screen.MousePointer = 11
    
    FechaDelServidor
        
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
        
    '27/03/2003 Nuevo campo CDiEliminado (fecha) de ahora en más se updatea -----------------------------
    Cons = "Select * from ChequeDiferido Where CDiCodigo = " & prmIdCheque
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.Edit
    RsAux!CDiEliminado = Format(gFechaServidor, "mm/dd/yyyy hh:mm")
    RsAux.Update: RsAux.Close
    '---------------------------------------------------------------------------------------------------------------------
    
    mTexto = tBanco.FormattedText & " " & Trim(tCSerie.Text) & "-" & Trim(tCNumero.Text)
    
    If Not bAlDia And Not myData.bRebotado Then 'Hago el ingreso de Caja    --> Si es DIFERIDO
        MovimientoDeCaja paMCChequeDiferido, gFechaServidor, mIDDisponibilidad, aIDMoneda, CCur(tCImporte.Text), "Dif. " & mTexto & " (eliminación)"
    End If
    
    'Registro el Suceso -------------------------------------------------------
    gSucesoDef = gSucesoDef & vbCrLf & zDatosCheque
    clsGeneral.RegistroSuceso cBase, gFechaServidor, prmSucesoCheque, paCodigoDeTerminal, gSucesoUsr, Val(lDoc.Tag), _
                            Descripcion:="Eliminación " & mTexto, _
                            Defensa:=Trim(gSucesoDef), _
                            Valor:=CCur(tCImporte.Text), idCliente:=(Val(lClienteC.Tag))
    
    cBase.CommitTrans    'Fin de la TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    LimpioFicha True
    Botones False, False, False, False, False, Toolbar1, Me
    
    Dim mID_Aux As Long: mID_Aux = prmIdCheque
    CargoCheque mID_Aux
    
    tBanco.SetFocus
    Screen.MousePointer = 0
    Exit Sub

errBEliminar:
    clsGeneral.OcurrioError "Error al validar los datos para eliminar.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
    
errorBT:
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación."
    Screen.MousePointer = 0: Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación."
End Sub

Private Function Zureo_Desdepositar(ByVal mTCDolar As Currency, ByVal mTCPesos As Currency) As Long
FechaDelServidor
    Screen.MousePointer = 11
    
    Dim RsMov As rdoResultset
    Dim mFechaHora As String
    Dim mIDMov As Long, mCompra As Long, I As Integer
    'Dim mTCDolar As Currency, mTCPesos As Currency
    
'    mTCPesos = 1
'    mTCDolar = TasadeCambio(paMonedaDolar, cMoneda.ItemData(cMoneda.ListIndex), UltimoDia(DateAdd("m", -1, gFechaServidor)))
'    If cMoneda.ItemData(cMoneda.ListIndex) <> paMonedaPesos Then
'        mTCPesos = TasadeCambio(cMoneda.ItemData(cMoneda.ListIndex), paMonedaPesos, UltimoDia(DateAdd("m", -1, gFechaServidor)))
'    End If
    mFechaHora = Format(gFechaServidor, "dd/mm/yyyy") & " " & Format(gFechaServidor, "hh:mm:ss")
    
    Dim objGeneric As New clsDBFncs
    Dim rdoCZureo As rdoConnection
    If Not objGeneric.get_Connection(rdoCZureo, "ORG01", 10) Then
        MsgBox "Error al conectarse a la base de datos de Zureo.", vbExclamation, "Conexión Zureo"
        Screen.MousePointer = 0
        Exit Function
    End If
    
    Dim m_ReturnID As Long
    Dim objComp As New clsComprobantes
    
    Dim OBJ_COM As clsDComprobante, OBJ_CTA As clsDCuenta
    Dim colCuentas As New Collection
    
    Dim xCuentaS As Long, xCuentaE As Long
    Dim xCuentaS_M As Integer, xCuentaE_M As Integer

    'Grabo los Gastos Y Transferencias  --------------------------------------------------------------------------
    For I = LBound(dGastos) To UBound(dGastos)
        mIDMov = 0: mCompra = 0
        
        If dGastos(I).ImporteAlDia <> 0 Then      'Transferencia entre disponibilidades
                    
            Cons = "Select DisID, DisIDSubRubro, DisMoneda  from Disponibilidad " & _
                   " Where DisID IN (" & dGastos(I).IdDisponibilidadEntrada & "," & dGastos(I).IdDisponibilidadSalida & ")"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            Do While Not RsAux.EOF
                If RsAux!DisID = dGastos(I).IdDisponibilidadEntrada Then
                    xCuentaE = RsAux!DisIDSubrubro
                    xCuentaE_M = RsAux!DisMoneda
                Else
                    xCuentaS = RsAux!DisIDSubrubro
                    xCuentaS_M = RsAux!DisMoneda
                End If
                RsAux.MoveNext
            Loop
            RsAux.Close

            Set OBJ_COM = New clsDComprobante
            With OBJ_COM
                .doAccion = 1 'IIf(xCta1 <> 0, 1, 0)
                .Ente = 0 'dGastos(I).IdProveedorGasto
                .Empresa = 1
                .Fecha = CDate(Format(mFechaHora, "dd/mm/yyyy"))
                .Tipo = 50  'Transferencias Zureo 'TipoDocumento.CompraEntradaCaja
                .Moneda = dGastos(I).idMoneda
                .ImporteTotal = dGastos(I).ImporteAlDia
                .TC = IIf(dGastos(I).idMoneda <> paMonedaPesos, mTCDolar, 1)
                .Memo = "Desdepósito de Ch.D. " & dGastos(I).SucursalNombre
                .UsuarioAlta = paCodigoDeUsuario
                .UsuarioAutoriza = paCodigoDeUsuario
            End With
        
            Set OBJ_CTA = New clsDCuenta
            With OBJ_CTA
                .VaAlDebe = 0
                .Cuenta = xCuentaS 'dGastos(I).IdSubrubroSalida
                .ImporteComp = dGastos(I).ImporteAlDia
                .ImporteCta = dGastos(I).ImporteAlDia * mTCPesos
                .MonedaCta = xCuentaS_M 'prmMonedaContabilidad
            End With
            colCuentas.Add OBJ_CTA
            Set OBJ_CTA = Nothing
            
            Set OBJ_CTA = New clsDCuenta
            With OBJ_CTA
                .VaAlDebe = 1
                .Cuenta = xCuentaE
                .ImporteComp = dGastos(I).ImporteAlDia
                .ImporteCta = dGastos(I).ImporteAlDia * mTCPesos
                .MonedaCta = xCuentaE_M 'prmMonedaContabilidad
            End With
            colCuentas.Add OBJ_CTA
            Set OBJ_CTA = Nothing
        
            Set OBJ_COM.Cuentas = colCuentas
            If objComp.fnc_PasarComprobante(rdoCZureo, OBJ_COM) Then m_ReturnID = objComp.prm_Comprobante
            Set OBJ_COM = Nothing
        End If
                
        mIDMov = 0
        If dGastos(I).ImporteDiferido <> 0 Then     'Registro Gasto por Cheques Diferidos
        
            Cons = "Select DisID, DisIDSubRubro, DisMoneda  from Disponibilidad " & _
                   " Where DisID = " & dGastos(I).IdDisponibilidadEntrada
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                xCuentaE = RsAux!DisIDSubrubro
                xCuentaE_M = RsAux!DisMoneda
                RsAux.MoveNext
            End If
            RsAux.Close
                               
            Set OBJ_COM = New clsDComprobante
            With OBJ_COM
                .doAccion = 1 'IIf(xCta1 <> 0, 1, 0)
                .Ente = 0 'dGastos(I).IdProveedorGasto
                .Empresa = 1
                .Fecha = CDate(Format(mFechaHora, "dd/mm/yyyy"))
                .Tipo = TipoDocumento.CompraEntradaCaja
                .Moneda = dGastos(I).idMoneda
                .ImporteTotal = dGastos(I).ImporteDiferido
                .TC = IIf(dGastos(I).idMoneda <> paMonedaPesos, mTCDolar, 1)
                .Memo = "Depósito de Ch.D. " & dGastos(I).SucursalNombre
                .UsuarioAlta = paCodigoDeUsuario
                .UsuarioAutoriza = paCodigoDeUsuario
            End With

        
            Set OBJ_CTA = New clsDCuenta
            With OBJ_CTA
                .VaAlDebe = 0
                .Cuenta = dGastos(I).IdSubrubroSalida
                .ImporteComp = dGastos(I).ImporteDiferido
                .ImporteCta = dGastos(I).ImporteDiferido * mTCPesos
                .MonedaCta = paMonedaPesos 'prmMonedaContabilidad
            End With
            colCuentas.Add OBJ_CTA
            Set OBJ_CTA = Nothing

            Set OBJ_CTA = New clsDCuenta
            With OBJ_CTA
                .VaAlDebe = 1
                .Cuenta = xCuentaE 'dGastos(I).IdDisponibilidadEntrada
                .ImporteComp = dGastos(I).ImporteDiferido
                .ImporteCta = dGastos(I).ImporteDiferido * mTCPesos
                .MonedaCta = xCuentaE_M 'prmMonedaContabilidad
            End With
            colCuentas.Add OBJ_CTA
            Set OBJ_CTA = Nothing

            Set OBJ_CTA = Nothing
            Set OBJ_COM.Cuentas = colCuentas
            
            If objComp.fnc_PasarComprobante(rdoCZureo, OBJ_COM) Then m_ReturnID = objComp.prm_Comprobante
            
            Set OBJ_COM = Nothing
        End If
    Next
    Set objComp = Nothing
    rdoCZureo.Close
    
    Zureo_Desdepositar = m_ReturnID
    
End Function

Private Sub AccionDesDepositar(bRebota As Boolean)

On Error GoTo errBEliminar

    Dim bAlDia As Boolean
    bAlDia = (myData.Vencimiento = CDate("1/1/1900"))
    
    If bRebota Then
        If MsgBox("Confirma sacar el cheque del Banco y darlo como Rebotado ?.", _
                         vbQuestion + vbYesNo + vbDefaultButton2, "Rebotar Cheque") = vbNo Then Exit Sub
    Else
        If MsgBox("Confirma sacar el cheque del Banco y dejarlo para depositar?.", _
                         vbQuestion + vbYesNo + vbDefaultButton2, "Sacar Cheque del Depósito") = vbNo Then Exit Sub

    End If
    
    If Not CargoDataGastos Then Exit Sub
    Me.Refresh
    
    On Error GoTo errorBT
    Screen.MousePointer = 11
    FechaDelServidor
    
    'Tasas para registrar mov. p/sacar cheque del Bco.  ---------------------------------------------------
    Dim mTCPesos As Currency, mTCDolar As Currency
    mTCPesos = 1
    If myData.Moneda <> paMonedaPesos Then
        mTCPesos = TasadeCambio(myData.Moneda, paMonedaPesos, UltimoDia(DateAdd("m", -1, gFechaServidor)))
    End If
    mTCDolar = TasadeCambio(paMonedaDolar, myData.Moneda, UltimoDia(DateAdd("m", -1, gFechaServidor)))
    
    Dim idZureo As Long
    idZureo = Zureo_Desdepositar(mTCDolar, mTCPesos)
    
    If idZureo = 0 Then
        MsgBox "No obtuve un ID de comprobante de Zureo, por favor reintente la operación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
        
    'Saco Datos de Depósito y Marco como Rebotado ----------------------------------------------------------
    Cons = "Select * from ChequeDiferido Where CDiCodigo = " & prmIdCheque
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    RsAux.Edit
    RsAux!CDiCobrado = Null
    RsAux!CDiDepositado = Null
    If bRebota Then RsAux!CDiRebotado = Format(gFechaServidor, "mm/dd/yyyy hh:mm")
    RsAux.Update: RsAux.Close
    '---------------------------------------------------------------------------------------------------------------------
    
    'Registro el Suceso -------------------------------------------------------
    mTexto = tBanco.FormattedText & " " & Trim(tCSerie.Text) & "-" & Trim(tCNumero.Text)
    gSucesoDef = zDatosCheque
    clsGeneral.RegistroSuceso cBase, gFechaServidor, prmSucesoCheque, paCodigoDeTerminal, paCodigoDeUsuario, Val(lDoc.Tag), _
                            Descripcion:="Elimino Depósito Ch. " & mTexto & " (Zureo: " & idZureo & ")", _
                            Defensa:=Trim(gSucesoDef), _
                            Valor:=CCur(tCImporte.Text), idCliente:=(Val(lClienteC.Tag))
    
    'FORMATO ANTERIOR (cambié a Zureo 9/11/2012
'    GraboMovimientosDesdeposito mTCDolar, mTCPesos
    
    
    cBase.CommitTrans    'Fin de la TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Screen.MousePointer = 0
    
    If idZureo > 0 Then
        MsgBox "Se creo el comprobante " & idZureo & " en Zureo, por favor valide que el mismo este correcto.", vbExclamation, "ATENCIÓN"
    End If
    
    LimpioFicha True
    Botones False, False, False, False, False, Toolbar1, Me
    
    Dim mID_Aux As Long: mID_Aux = prmIdCheque
    CargoCheque mID_Aux
    
    tBanco.SetFocus
    Exit Sub

errBEliminar:
    clsGeneral.OcurrioError "Error al validar los datos para ejecutar el procedimiento.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
    
errorBT:
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación."
    Screen.MousePointer = 0: Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación."
End Sub

Private Sub AccionModificar()

    If prmIdCheque = 0 Then
        MsgBox "Debe seleccionar un cheque para modificar los datos.", vbExclamation, "Falta Seleccionar Cheque"
        tBanco.SetFocus: Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    CargoCheque prmIdCheque
    If prmIdCheque = 0 Then Exit Sub
    
    sModificar = True
    Botones False, False, False, True, True, Toolbar1, Me
    EstadoIngreso True
    Foco tCSerie
    
    Screen.MousePointer = 0
    
End Sub

Private Sub AccionCancelar()

On Error Resume Next

    Screen.MousePointer = 11
    
    EstadoIngreso False
    sModificar = False
    LimpioFicha
    Botones False, False, False, False, False, Toolbar1, Me
    
    If prmIdCheque <> 0 Then CargoCheque prmIdCheque
    tBanco.SetFocus
    
    Screen.MousePointer = 0
    
End Sub

Private Sub EstadoIngreso(bEstado As Boolean)

Dim bkColor As Long
    
    If bEstado Then bkColor = Colores.Blanco Else bkColor = Colores.Inactivo
    
    If (bEstado And Not myData.bDepositado) Or Not bEstado Then
        tCVencimiento.Enabled = bEstado: tCVencimiento.BackColor = bkColor
        tCImporte.Enabled = bEstado: tCImporte.BackColor = bkColor
        tCLibrado.Enabled = bEstado: tCLibrado.BackColor = bkColor
        tDocumento.Enabled = bEstado: tDocumento.BackColor = bkColor
    End If
    
    tComentario.Enabled = bEstado: tComentario.BackColor = bkColor
            
End Sub

Private Sub AccionGrabar()
'   -->    Graba Modificaciones, si estaba depositado solo modifica nros de cheque y comentarios

On Error GoTo errBGrabar

Dim aUsuario As Long
Dim aDiferencia As Currency
Dim mIDDisponibilidad As Long, aIDMoneda As Integer

    If Not ValidoDatos Then Exit Sub
    
Dim bAlDia As Boolean

    bAlDia = (myData.Vencimiento = CDate("1/1/1900"))
    
    aDiferencia = myData.Importe - CCur(tCImporte.Text)     'Lo q' Habia - lo nuevo
    
    If aDiferencia = 0 Or bAlDia Then
        If MsgBox("Confirma almacenar los datos ingresados.", vbQuestion + vbYesNo, "Grabar Datos") = vbNo Then Exit Sub
    
    Else        'Para cheques Diferidos con Diferencia en Importe
    
        If aDiferencia > 0 Then mTexto = " pesos de menos, " Else mTexto = " pesos de más, "
        mTexto = "Hay una diferencia de " & Format(Abs(aDiferencia), FormatoMonedaP) & mTexto & "contra el importe anterior." & vbCrLf _
                & "El importe era de " & Format(myData.Importe, FormatoMonedaP) & " y pasó a " & Format(tCImporte.Text, FormatoMonedaP) & "." & vbCrLf & vbCrLf _
                & "Si = Realizar un movimiento de caja para compensar la diferencia." & vbCrLf _
                & "No = Cancelar la modificación."
                
        If MsgBox(mTexto, vbQuestion + vbYesNo + vbDefaultButton2, "Grabar Datos") = vbNo Then Exit Sub
    
        aIDMoneda = Val(myData.Moneda)   'Moneda
        mIDDisponibilidad = dis_DisponibilidadPara(paCodigoDeSucursal, CLng(aIDMoneda))
        If mIDDisponibilidad = 0 Then
            MsgBox "Su sucursal no puede recibir un cheque en " & Trim(lMoneda.Caption) & "." & vbCrLf & _
                        "No hay una disponibilidad para realizar los movimientos de caja que correspondan." & vbCrLf & vbCrLf & _
                        "Consulte con el administrador del sistema.", vbExclamation, "Falta Disponibilidad para la Moneda del Cheque"
            Screen.MousePointer = 0
            Exit Sub
        End If
        
    End If
    
    FechaDelServidor
    
    'Llamo al registro del Suceso-------------------------------------------------------------
    Dim objSuceso As New clsSuceso
    objSuceso.ActivoFormulario paCodigoDeUsuario, "Modificación de Cheques Diferidos", cBase
    Me.Refresh
    gSucesoUsr = objSuceso.RetornoValor(Usuario:=True)
    gSucesoDef = objSuceso.RetornoValor(Defensa:=True)
    Set objSuceso = Nothing
    If gSucesoUsr = 0 Then Screen.MousePointer = 0: Exit Sub
    '---------------------------------------------------------------------------------------------
    
    FechaDelServidor
    
    On Error GoTo errorBT
    Screen.MousePointer = 11
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
    
    Cons = "Select * from ChequeDiferido Where CDiCodigo = " & prmIdCheque
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.Edit
    RsAux!CDiBanco = zBancoDelTexto(tBanco.Tag)
    RsAux!CDiSucursal = zSucursalDelTexto(tBanco.Tag)
    RsAux!CDiSerie = Trim(tCSerie.Text)
    RsAux!CDiNumero = tCNumero.Text
    
    If Not bAlDia Then RsAux!CDiVencimiento = Format(tCVencimiento.Text, "mm/dd/yyyy") Else RsAux!CDiVencimiento = Null
    
    RsAux!CDiLibrado = Format(tCLibrado.Text, "mm/dd/yyyy")
    RsAux!CDiUsuario = aUsuario
    RsAux!CDiImporte = CCur(tCImporte.Text)
    If Trim(tComentario.Text) <> "" Then RsAux!CDiComentario = Trim(tComentario.Text) Else RsAux!CDiComentario = Null
    
    If Val(lDoc.Tag) <> 0 Then RsAux!CDiDocumento = Val(lDoc.Tag) Else RsAux!CDiDocumento = Null
    
    RsAux.Update: RsAux.Close

    'Registro el Suceso -----------------------------------------------------------------------------------------------------------
    mTexto = tBanco.FormattedText & " " & Trim(tCSerie.Text) & "-" & Trim(tCNumero.Text)
    If mTexto <> myData.NumeroCheque Then mTexto = mTexto & " (" & myData.NumeroCheque & ")"
    gSucesoDef = zVeoCambios(gSucesoDef)
    
    clsGeneral.RegistroSuceso cBase, gFechaServidor, prmSucesoCheque, paCodigoDeTerminal, gSucesoUsr, Val(lDoc.Tag), _
                            Descripcion:="Modificación " & mTexto, _
                            Defensa:=Trim(gSucesoDef), Valor:=aDiferencia, idCliente:=(Val(lClienteC.Tag))
    
    'Hago el movimiento de caja por la diferencia
    If aDiferencia <> 0 And Not bAlDia Then
        Dim bSalida As Boolean: If aDiferencia > 0 Then bSalida = False Else bSalida = True
        mTexto = "Dif. " & tBanco.FormattedText & " " & Trim(tCSerie.Text) & "-" & Trim(tCNumero.Text) & " (diferencia en importe)"
        
        MovimientoDeCaja paMCChequeDiferido, gFechaServidor, mIDDisponibilidad, aIDMoneda, Abs(aDiferencia), mTexto, bSalida
                           
    End If
        
    cBase.CommitTrans    'Fin de la TRANSACCION ------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    sModificar = False
    EstadoBotones
    EstadoIngreso False
    Screen.MousePointer = 0
    Exit Sub

errBGrabar:
    clsGeneral.OcurrioError "Error al validar los datos para grabar.", Err.Description
    Screen.MousePointer = 0: Exit Sub
    
errorBT:
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación."
    Screen.MousePointer = 0: Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación."
    Screen.MousePointer = 0
End Sub

Private Sub LimpioFicha(Optional Todo As Boolean = False)
    
    If Todo Then
        tBanco.Text = ""
        lBanco.Caption = "S/D"
        tCSerie.Text = ""
        tCNumero.Text = ""
    End If
    
    lClienteC.Caption = ""
    lClienteD.Caption = ""
    
    tCVencimiento.Text = ""
    tCLibrado.Text = ""
    tCImporte.Text = ""
    tComentario.Text = ""
    tDocumento.Text = "": lDoc.Caption = ""
    
    lUsuario.Caption = ""
    lMoneda.Caption = ""
    lFDepositado.Caption = ""
    lFIngresado.Caption = ""
    lBancoDeposito.Caption = ""
    lEliminado.Caption = "": lEliminado.Visible = False: lEliminadoT.Visible = False 'lEliminado.BackColor = lFIngresado.BackColor
    lRebotado.Caption = "": lRebotado.Visible = False: lRebotadoT.Visible = False 'lRebotado.BackColor = lFIngresado.BackColor
    
End Sub

Private Sub BuscoCheque()
    
    If Trim(tCSerie.Text) = "" Then
        MsgBox "Los datos del cheque están incompletos.", vbExclamation, "Faltan Datos"
        Foco tCSerie: Exit Sub
    End If
    If Not IsNumeric(tCNumero.Text) Then
        MsgBox "Los datos del cheque están incompletos.", vbExclamation, "Faltan Datos"
        Foco tCNumero: Exit Sub
    End If
    
    On Error GoTo errCargar
    prmIdCliente = 0: prmIdCheque = 0

    Cons = " Select CDiCodigo, CDiSerie as Serie, CDiNumero as Numero, BanNombre as Banco, SBaNombre as Sucursal " & _
                " From ChequeDiferido, SucursalDeBanco, BancoSSFF " & _
                " Where CDiSerie = '" & Trim(tCSerie.Text) & "'" & _
                " And CDiNumero = " & tCNumero.Text & _
                " And CDiSucursal = SBaCodigo" & _
                " And CDiBanco = BanCodigo"
    
    If Val(tBanco.Tag) <> 0 Then
        Cons = Cons & " And CDiBanco = " & zBancoDelTexto(tBanco.Tag) & " And CDiSucursal = " & zSucursalDelTexto(tBanco.Tag)
    End If

    Dim aQ As Integer, aIdSel As Long
    aQ = 0: aIdSel = 0
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        aQ = 1
        aIdSel = RsAux!CDiCodigo
        RsAux.MoveNext: If Not RsAux.EOF Then aQ = 2
    End If
    RsAux.Close
    
    Select Case aQ
        Case 0: MsgBox "No hay datos que coincidan con los valores ingresados.", vbExclamation, "No hay datos"
        
        Case 2:
                    Dim miLista As New clsListadeAyuda
                    aIdSel = miLista.ActivarAyuda(cBase, Cons, 6000, 1, "Lista de Cheques")
                    Me.Refresh
                    If aIdSel > 0 Then aIdSel = miLista.RetornoDatoSeleccionado(0)
                    Set miLista = Nothing
    End Select
    
    If aIdSel > 0 Then CargoCheque aIdSel
    Screen.MousePointer = 0
    
    Exit Sub
errCargar:
    clsGeneral.OcurrioError "Error al buscar el cheque.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoCheque(idCheque As Long)
    
    On Error GoTo errCargar
    Screen.MousePointer = 11
    LimpioFicha
    
    myData.Importe = 0
    myData.Moneda = 0
    myData.NumeroCheque = ""
    myData.Documento = 0
    myData.Vencimiento = CDate("1/1/1900")
    myData.bDepositado = False
    myData.bRebotado = False
    myData.IdSucursalDepositado = 0
    
    Cons = "Select * from ChequeDiferido, SucursalDeBanco, BancoSSFF " _
            & " Where CDiCodigo = " & idCheque _
            & " And CDiSucursal = SBaCodigo" _
            & " And CDiBanco = BanCodigo"
            
    prmIdCheque = 0
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        prmIdCheque = RsAux!CDiCodigo
        prmIdCliente = RsAux!CDiCliente
                
        lClienteC.Tag = RsAux!CDiCliente
        lClienteD.Tag = RsAux!CDiClienteFactura
        
        tBanco.Text = Format(RsAux!BanCodigoB, "00") & Format(RsAux!SBaCodigoS, "000")
        tBanco.Tag = RsAux!BanCodigo & "-" & RsAux!SBaCodigo
        lBanco.Caption = Trim(RsAux!BanNombre) & " - " & Trim(RsAux!SBaNombre)
        
        With myData
            .Importe = Format(RsAux!CDiImporte, FormatoMonedaP)
            .Moneda = RsAux!CDiMoneda
            If Not IsNull(RsAux!CDiVencimiento) Then .Vencimiento = RsAux!CDiVencimiento
            If Not IsNull(RsAux!CDiDocumento) Then .Documento = RsAux!CDiDocumento
            .NumeroCheque = tBanco.FormattedText & " " & Trim(RsAux!CDiSerie) & "-" & RsAux!CDiNumero
            .Librado = RsAux!CDiLibrado
            If Not IsNull(RsAux!CDiDepositado) Then .bDepositado = True
        End With
        
        tCSerie.Text = Trim(RsAux!CDiSerie)
        tCNumero.Text = RsAux!CDiNumero
        tCImporte.Text = Format(RsAux!CDiImporte, FormatoMonedaP)
        If Not IsNull(RsAux!CDiVencimiento) Then tCVencimiento.Text = Format(RsAux!CDiVencimiento, "dd/mm/yyyy")
        tCLibrado.Text = Format(RsAux!CDiLibrado, "dd/mm/yyyy")
        
        If Not IsNull(RsAux!CDiComentario) Then tComentario.Text = Trim(RsAux!CDiComentario)
        
        If Not IsNull(RsAux!CDiRebotado) Then
            myData.bRebotado = True
            lRebotado.Caption = " " & Format(RsAux!CDiRebotado, "dd/mm/yyyy hh:mm")
            lRebotado.BackColor = Colores.Rojo: lRebotado.Visible = True: lRebotadoT.Visible = True
        End If
        
        lUsuario.Caption = " " & zBuscoNombreUsuario(RsAux!CDiUsuario)
        lFIngresado.Caption = " " & Format(RsAux!CDiIngresado, "dd/mm/yyyy hh:mm")
        lMoneda.Caption = " " & zBuscoNombreMoneda(RsAux!CDiMoneda)
    
        If Not IsNull(RsAux!CDiCobrado) Then lFDepositado.Caption = " " & Format(RsAux!CDiCobrado, "dd/mm/yyyy hh:mm")
        If Not IsNull(RsAux!CDiDepositado) Then
            
            myData.IdSucursalDepositado = RsAux!CDiDepositado
            
            Cons = "Select * from SucursalDeBanco, BancoSSFF" _
                    & " Where SBaCodigo = " & RsAux!CDiDepositado _
                    & " And SBaBanco = BanCodigo"
            Set rsX = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not rsX.EOF Then lBancoDeposito.Caption = " " & Trim(rsX!BanNombre) & " - " & Trim(rsX!SBaNombre)
            rsX.Close
        End If
    
        If Not IsNull(RsAux!CDiEliminado) Then
            lEliminado.Caption = " " & Format(RsAux!CDiEliminado, "dd/mm/yyyy hh:mm")
            lEliminado.BackColor = Colores.Rojo: lEliminado.Visible = True: lEliminadoT.Visible = True
        End If
        
    Else
        Screen.MousePointer = 0
        MsgBox "El cheque ingresado no existe.", vbInformation, "No hay Datos"
    End If
    RsAux.Close
    
    EstadoBotones
    
    If prmIdCheque <> 0 Then
        lClienteC.Caption = zNombreCliente(CLng(lClienteC.Tag))
        lClienteD.Caption = zNombreCliente(CLng(lClienteD.Tag))
        If myData.Documento <> 0 Then zBuscoDocumento myData.Documento
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Error al cargar los datos del cheque.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub EstadoBotones()
    
    If prmIdCheque = 0 Or Trim(lEliminado.Caption) <> "" Then
        Botones False, False, False, False, False, Toolbar1, Me
        With Toolbar1
            .Buttons("rebotar").Enabled = False
            .Buttons("desdepositar").Enabled = False
        End With
        
        Exit Sub
    End If
        
    Dim bAlDia As Boolean
    bAlDia = (myData.Vencimiento = CDate("1/1/1900"))
    
    Dim bBRebDes As Boolean
    bBRebDes = False
    
    If myData.bDepositado Then
        Botones False, True, False, False, False, Toolbar1, Me
        If Not myData.bRebotado Then bBRebDes = True
    Else
        Botones Not bAlDia, True, True, False, False, Toolbar1, Me
    End If
    
    With Toolbar1
        .Buttons("rebotar").Enabled = bBRebDes
        .Buttons("desdepositar").Enabled = bBRebDes
    End With
        
End Sub

Private Function zBancoDelTexto(Texto As String) As Long
    On Error Resume Next
    zBancoDelTexto = 0
    zBancoDelTexto = CLng(Mid(Texto, 1, InStr(Texto, "-") - 1))
End Function

Private Function zSucursalDelTexto(Texto As String) As Long
    On Error Resume Next
    zSucursalDelTexto = 0
    zSucursalDelTexto = CLng(Mid(Texto, InStr(Texto, "-") + 1, Len(Texto)))
End Function

Private Function ValidoDatos() As Boolean

    On Error GoTo errValido
    ValidoDatos = False
    
    If tBanco.Tag = "" Then
        MsgBox "Ingrese el código de banco emisor que figura en el cheque.", vbExclamation, "Falta Bco Emisor"
        tBanco.SetFocus: Exit Function
    End If
    
    If Trim(tCSerie.Text) = "" Then
        MsgBox "La serie del cheque ingresada no es correcta.", vbExclamation, "Faltan Datos del Cheque"
        Foco tCSerie: Exit Function
    End If
    If Not IsNumeric(tCNumero.Text) Or Trim(tCNumero.Text) = "" Then
        MsgBox "El número de cheque ingresado no es correcto.", vbExclamation, "Faltan Datos del Cheque"
        Foco tCNumero: Exit Function
    End If
    
    If myData.Vencimiento <> CDate("1/1/1900") Then     'ERA DIFERIDO
        If Not IsDate(tCVencimiento.Text) Then
            MsgBox "La fecha de vencimiento ingresada no es correcta.", vbExclamation, "Faltan Datos del Cheque"
            Foco tCVencimiento: Exit Function
        End If
    Else
        If Trim(tCVencimiento.Text) <> "" Then
            MsgBox "El cheque era un cheque al día, no se puede modificar como diferido.", vbExclamation, "Cheque Al Día"
            Foco tCVencimiento: Exit Function
        End If
    End If
    
    If Not IsDate(tCLibrado.Text) Then
        MsgBox "La fecha de librado ingresada no es correcta.", vbExclamation, "Faltan Datos del Cheque"
        Foco tCLibrado: Exit Function
    End If
    
    If Not IsNumeric(tCImporte.Text) Then
        MsgBox "El importe del cheque no es correcto. Verifique.", vbExclamation, "Faltan Datos del Cheque"
        Foco tCImporte: Exit Function
    End If
    ValidoDatos = True
    Exit Function
    
errValido:
    clsGeneral.OcurrioError "Error al validar los datos. " & Err.Description
    Screen.MousePointer = 0
End Function

Function zBuscoNombreMoneda(Codigo As Long) As String
On Error GoTo ErrBU
    zBuscoNombreMoneda = ""

    Cons = "Select * From Moneda Where MonCodigo = " & Codigo
    Set rsX = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsX.EOF Then zBuscoNombreMoneda = Trim(rsX!MonNombre)
    rsX.Close
ErrBU:
End Function

Function zBuscoNombreUsuario(Codigo As Long) As String
On Error GoTo ErrBU
    
    zBuscoNombreUsuario = ""
    Cons = "Select * From Usuario Where UsuCodigo = " & Codigo
    Set rsX = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsX.EOF Then zBuscoNombreUsuario = Trim(rsX!UsuIdentificacion)
    rsX.Close
    
ErrBU:
End Function

Private Sub zBuscoDocumento(mIDDoc As Long)
On Error GoTo errBDoc
Dim adCodigo As Long, adTexto As String, adTextoFmt As String

    Cons = "Select * from Documento " & _
               " Where DocCodigo = " & mIDDoc
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        adCodigo = RsAux!DocCodigo
        adTexto = zDocumento(RsAux!DocTipo, RsAux!DocSerie, RsAux!DocNumero, adTextoFmt)
    End If
    RsAux.Close
    
    If adCodigo > 0 Then
        tDocumento.Text = adTextoFmt
        lDoc.Tag = adCodigo: lDoc.Caption = adTexto
    Else
        lDoc.Caption = " No Existe !!"
    End If
    Exit Sub
    
errBDoc:
    clsGeneral.OcurrioError "Error al buscar el documento.", Err.Description
End Sub

Private Function zVeoCambios(mCom As String) As String
On Error Resume Next
Dim mRet As String: mRet = ""

    With myData
        If .Vencimiento <> CDate(tCVencimiento.Text) And Trim(tCVencimiento.Text) <> "" Then
            mRet = "Vto.An.: " & Format(.Vencimiento, "dd/mm/yy")
        End If
        
        If .Importe <> CCur(tCImporte.Text) Then
            If mRet <> "" Then mRet = mRet & ", "
            mRet = mRet & "$.An.: " & Format(.Importe, "##0.00")
        End If
        
        If .Librado <> CDate(tCLibrado.Text) Then
            If mRet <> "" Then mRet = mRet & ", "
            mRet = mRet & "F/Lib.An.: " & Format(.Librado, "dd/mm/yy")
        End If
        
    End With
    
    If mRet <> "" Then zVeoCambios = mCom & vbCrLf & mRet Else zVeoCambios = mCom
    
End Function

Private Function zDatosCheque() As String
On Error Resume Next
Dim mRet As String: mRet = ""

    With myData
        If Trim(tCVencimiento.Text) <> "" Then
            mRet = "Vto: " & Format(.Vencimiento, "dd/mm/yy")
        End If
        
        If mRet <> "" Then mRet = mRet & ", "
        mRet = mRet & "F/Lib.: " & Format(.Librado, "dd/mm/yy") & ", "
        If myData.bRebotado Then mRet = mRet & "Reb." Else mRet = mRet & "No Reb."
    End With
    
    zDatosCheque = mRet
    
End Function

Private Function CargoDataGastos() As Boolean

Dim I As Integer, bOK As Boolean
    
    CargoDataGastos = False
    
    ReDim dGastos(0)

    bOK = arrG_AddItem(myData.IdSucursalDepositado, lBancoDeposito.Caption, myData.Importe, myData.Vencimiento = CDate("1/1/1900"), CLng(myData.Moneda), 0)
            
    If Not bOK Then
        clsGeneral.OcurrioError "Errores al procesar el cheque.", Err.Description
        Exit Function
    End If
    
    frmWizGasto.Show vbModal, Me
    If Not frmWizGasto.prmOK Then Exit Function
    
    CargoDataGastos = True
    
End Function

'--------------------------------------------------------------------------------------------------------------------------------------
'   Cuando el Cheque está depositado hay que grabar un movimiento (Gasto o Trasferencia) para sacarlo
'   del Banco. Lo inverso a lo que se hace en el depósito.
    
Private Function GraboMovimientosDesdeposito(mTCDolar As Currency, mTCPesos As Currency)
'Tasas de la moneda del cheque a posos/dolares

Dim RsMov As rdoResultset
Dim mFechaHora As String
Dim mIDMov As Long, mCompra As Long, I As Integer
    
    mIDMov = 0: mCompra = 0
    mFechaHora = Format(gFechaServidor, "dd/mm/yyyy") & " " & Format(gFechaServidor, "hh:mm:ss")
    
    If dGastos(0).ImporteAlDia <> 0 Then      'Transferencia entre disponibilidades
        'Inserto en la Tabla Movimiento-Disponibilidad--------------------------------------------------------
        Cons = "Select * from MovimientoDisponibilidad Where MDiID = " & mIDMov
        Set RsMov = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        RsMov.AddNew
        RsMov!MDiFecha = Format(mFechaHora, "mm/dd/yyyy")
        RsMov!MDiHora = Format(mFechaHora, "hh:mm:ss")
        RsMov!MDiTipo = dGastos(0).IdTipoTransferencia
        RsMov!MDiIdCompra = Null
        RsMov!MDiComentario = "Elimino Depósito Ch." & myData.NumeroCheque & " (" & dGastos(0).SucursalNombre & ")"
        RsMov.Update: RsMov.Close
        '------------------------------------------------------------------------------------------------------------

        'Saco el Id de movimiento-------------------------------------------------------------------------------
        Cons = "Select Max(MDiID) from MovimientoDisponibilidad" & _
                  " Where MDiFecha = " & Format(mFechaHora, "'mm/dd/yyyy'") & _
                  " And MDiHora = " & Format(mFechaHora, "'hh:mm:ss'") & _
                  " And MDiTipo = " & dGastos(0).IdTipoTransferencia & _
                  " And MDiIdCompra is Null"
    
        Set RsMov = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        mIDMov = RsMov(0)
        RsMov.Close
        '------------------------------------------------------------------------------------------------------------

        'Grabo en Tabla Movimiento-Disponibilidad-Renglon--------------------------------------------------
        Cons = "Select * from MovimientoDisponibilidadRenglon Where MDRIdMovimiento = " & mIDMov
        Set RsMov = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        RsMov.AddNew        'Salida
        RsMov!MDRIdMovimiento = mIDMov
        RsMov!MDRIdDisponibilidad = dGastos(0).IdDisponibilidadSalida
        RsMov!MDRIdCheque = 0
        RsMov!MDRImporteCompra = dGastos(0).ImporteAlDia
        RsMov!MDRImportePesos = dGastos(0).ImporteAlDia * mTCPesos
        RsMov!MDRHaber = dGastos(0).ImporteAlDia
        RsMov.Update
        
        RsMov.AddNew        'Entrada
        RsMov!MDRIdMovimiento = mIDMov
        RsMov!MDRIdDisponibilidad = dGastos(0).IdDisponibilidadEntrada
        RsMov!MDRIdCheque = 0
        RsMov!MDRImporteCompra = dGastos(0).ImporteAlDia
        RsMov!MDRImportePesos = dGastos(0).ImporteAlDia * mTCPesos
        RsMov!MDRDebe = dGastos(0).ImporteAlDia
        RsMov.Update
        
        RsMov.Close
    End If
                
    mIDMov = 0
    If dGastos(0).ImporteDiferido <> 0 Then     'Registro Gasto por Cheques Diferidos
    
        Cons = "Select * from Compra Where ComCodigo = " & mCompra
        Set RsMov = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsMov.AddNew
        
        RsMov!ComFecha = Format(mFechaHora, "mm/dd/yyyy")
        RsMov!ComProveedor = dGastos(0).IdProveedorGasto
        RsMov!ComTipoDocumento = TipoDocumento.CompraSalidaCaja
        
        RsMov!ComMoneda = dGastos(0).idMoneda
        RsMov!ComTC = mTCDolar
        RsMov!ComImporte = dGastos(0).ImporteDiferido
        RsMov!ComSaldo = 0
        
        RsMov!ComComentario = "Elimino Depósito Ch." & myData.NumeroCheque & " (" & dGastos(0).SucursalNombre & ")"
        RsMov!ComFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
        RsMov!ComUsuario = paCodigoDeUsuario
        RsMov.Update: RsMov.Close
        '--------------------------------------------------------------------------------------------------------------------------------

        Cons = "Select Max(ComCodigo) from Compra" & _
                " Where ComFecha = " & Format(mFechaHora, "'mm/dd/yyyy'") & _
                " And ComTipoDocumento = " & TipoDocumento.CompraSalidaCaja & _
                " And ComProveedor = " & dGastos(0).IdProveedorGasto & _
                " And ComMoneda = " & dGastos(0).idMoneda
        Set RsMov = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        mCompra = RsMov(0)
        RsMov.Close

        'Tabla Gasto Subrubros  ----------------------------------------------------------------------
        Cons = "Select * from GastoSubrubro Where GSrIDCompra = " & mCompra
        Set RsMov = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsMov.AddNew
        RsMov!GSrIDCompra = mCompra
        RsMov!GSrIDSubrubro = dGastos(0).IdSubrubroSalida
        RsMov!GSrImporte = dGastos(0).ImporteDiferido
        RsMov.Update: RsMov.Close

        'Inserto en la Tabla Movimiento-Disponibilidad--------------------------------------------------------
        Cons = "Select * from MovimientoDisponibilidad Where MDiID = " & mIDMov
        Set RsMov = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        RsMov.AddNew
        RsMov!MDiFecha = Format(mFechaHora, "mm/dd/yyyy")
        RsMov!MDiHora = Format(mFechaHora, "hh:mm:ss")
        RsMov!MDiTipo = paMDPagoDeCompra
        RsMov!MDiIdCompra = mCompra
        RsMov!MDiComentario = "Elimino Depósito Ch." & myData.NumeroCheque & " (" & dGastos(0).SucursalNombre & ")"
        RsMov.Update: RsMov.Close
        '------------------------------------------------------------------------------------------------------------
        
        'Saco el Id de movimiento-------------------------------------------------------------------------------
        Cons = "Select Max(MDiID) from MovimientoDisponibilidad" & _
                  " Where MDiFecha = " & Format(mFechaHora, "'mm/dd/yyyy'") & _
                  " And MDiHora = " & Format(mFechaHora, "'hh:mm:ss'") & _
                  " And MDiTipo = " & paMDPagoDeCompra & _
                  " And MDiIdCompra = " & mCompra
        Set RsMov = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        mIDMov = RsMov(0)
        RsMov.Close
        '------------------------------------------------------------------------------------------------------------
        
        'Grabo en Tabla Movimiento-Disponibilidad-Renglon--------------------------------------------------
        Cons = "Select * from MovimientoDisponibilidadRenglon Where MDRIdMovimiento = " & mIDMov
        Set RsMov = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If RsMov.EOF Then RsMov.AddNew Else RsMov.Edit
        
        RsMov!MDRIdMovimiento = mIDMov
        RsMov!MDRIdDisponibilidad = dGastos(0).IdDisponibilidadEntrada
        RsMov!MDRIdCheque = 0
        
        RsMov!MDRImporteCompra = dGastos(0).ImporteDiferido
        RsMov!MDRImportePesos = dGastos(0).ImporteDiferido * mTCPesos
        RsMov!MDRHaber = dGastos(0).ImporteDiferido
        RsMov.Update: RsMov.Close
    End If

End Function
