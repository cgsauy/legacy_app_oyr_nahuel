VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCedoProducto 
   Caption         =   "Ceder Producto"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCedoProducto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   5355
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fCliente1 
      Caption         =   "Datos Cliente"
      Height          =   2595
      Left            =   60
      TabIndex        =   16
      Top             =   2700
      Width           =   7275
      Begin MSMask.MaskEdBox tCI1 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox tRUC1 
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
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
      Begin VSFlex6DAOCtl.vsFlexGrid vsProducto1 
         Height          =   1155
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   2037
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   4
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0   'False
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Pro&ductos:"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lTitular1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rodriguez Fernandez, Rodrigo Bernardino"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3720
         TabIndex        =   21
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   3375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Dirección:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   570
         Width           =   705
      End
      Begin VB.Label lDireccion1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Niagara 2345"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   960
         TabIndex        =   19
         Top             =   570
         UseMnemonic     =   0   'False
         Width           =   5955
      End
      Begin VB.Label lTelefono1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Casa 9242557; Celular 099405236"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   825
         UseMnemonic     =   0   'False
         Width           =   6015
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfonos:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   825
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "C.I./&RUC:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   255
         Width           =   855
      End
   End
   Begin VB.Frame fCliente 
      Caption         =   "Datos Cliente"
      Height          =   2535
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   7275
      Begin MSMask.MaskEdBox tCi 
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox tRuc 
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
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
      Begin VSFlex6DAOCtl.vsFlexGrid vsProducto 
         Height          =   1095
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   1931
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   4
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0   'False
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Productos:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&C.I./RUC:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label56 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfonos:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   825
         Width           =   855
      End
      Begin VB.Label lTelefono 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Casa 9242557; Celular 099405236"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   825
         UseMnemonic     =   0   'False
         Width           =   6015
      End
      Begin VB.Label lDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Niagara 2345"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   570
         UseMnemonic     =   0   'False
         Width           =   5955
      End
      Begin VB.Label lNDireccion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Dirección:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   570
         Width           =   705
      End
      Begin VB.Label lTitular 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rodriguez Fernandez, Rodrigo Bernardino"
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3720
         TabIndex        =   11
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   3435
      End
   End
   Begin VB.Menu MnuBuscar 
      Caption         =   "Buscar"
      Visible         =   0   'False
      Begin VB.Menu MnuBuNuevo 
         Caption         =   "Nuevo"
         Checked         =   -1  'True
         Shortcut        =   {F3}
      End
      Begin VB.Menu MnuBuNuevoCliente 
         Caption         =   "Nuevo Cliente"
      End
      Begin VB.Menu MnuBuNuevaEmpresa 
         Caption         =   "Nueva Empresa"
      End
      Begin VB.Menu MnuBuLinea2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBusquedas 
         Caption         =   "Búsquedas"
         Shortcut        =   {F4}
      End
      Begin VB.Menu MnuBuBuscarPersonas 
         Caption         =   "Buscar Personas"
      End
      Begin VB.Menu MnuBuBuscarEmpresas 
         Caption         =   "Buscar Empresas"
      End
   End
End
Attribute VB_Name = "frmCedoProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gCliente As Long

Private Sub Form_Activate()
    Screen.MousePointer = 0: Me.Refresh
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyX: Unload Me
        End Select
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    ObtengoSeteoForm Me
    FechaDelServidor
    InicializoGrillaProducto
    LimpioFichaCliente
    LimpioFichaCliente1
    If Trim(Command()) <> "" Then CargoDatosClienteDado CLng(Command())
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrio un error al iniciar el formulario.", Trim(Err.Description)
End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    fCliente.Height = (Me.ScaleHeight - (Status.Height + fCliente.Top + 40)) / 2
    fCliente1.Top = fCliente.Top + fCliente.Height + 40
    fCliente1.Height = fCliente.Height - 80
    fCliente.Width = Me.ScaleWidth - (fCliente.Left * 2)
    fCliente1.Width = fCliente.Width
    
    lTitular.Width = fCliente.Width - (lTitular.Left + 50)
    lTitular1.Width = fCliente1.Width - (lTitular1.Left + 50)
    
    vsProducto.Width = fCliente.Width - (vsProducto.Left * 2)
    vsProducto1.Width = vsProducto.Width
    vsProducto.Height = fCliente.Height - (vsProducto.Top + 60)
    vsProducto1.Height = fCliente1.Height - (vsProducto1.Top + 60)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    GuardoSeteoForm Me
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
End Sub

Private Sub InicializoGrillaProducto()
    With vsProducto
        .Rows = 1
        .Cols = 1
        .ExtendLastCol = True
        .FormatString = "ID|Tipo de Artículo|Estado|>F.Compra|Garantía|N° Serie|Factura|"
        .ColWidth(0) = 650: .ColWidth(1) = 3000: .ColWidth(3) = 1000: .ColWidth(5) = 1100
    End With
    With vsProducto1
        .Rows = 1
        .Cols = 1
        .ExtendLastCol = True
        .FormatString = "ID|Tipo de Artículo|Estado|>F.Compra|Garantía|N° Serie|Factura|"
        .ColWidth(0) = 650: .ColWidth(1) = 3000: .ColWidth(3) = 1000: .ColWidth(5) = 1100
    End With
End Sub

Private Sub MnuBuBuscarEmpresas_Click()
    BuscarClientes TipoCliente.Empresa
End Sub

Private Sub MnuBuBuscarPersonas_Click()
    BuscarClientes TipoCliente.Cliente
End Sub

Private Sub MnuBuNuevaEmpresa_Click()
    NuevoCliente TipoCliente.Empresa
End Sub

Private Sub MnuBuNuevoCliente_Click()
    NuevoCliente TipoCliente.Cliente
End Sub

Private Sub tCi_Change()
    tCi.Tag = "0"
End Sub

Private Sub tCi_GotFocus()
    tCi.SelStart = 0: tCi.SelLength = 11
End Sub

Private Sub tCi_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 93: PopupMenu MnuBuscar, , tCi.Left + (tCi.Width / 2), (tCi.Top + tCi.Height) - (tCi.Height / 2)
        Case vbKeyF3: If Shift = 0 Then NuevoCliente TipoCliente.Cliente
        Case vbKeyF4: If Shift = 0 Then BuscarClientes TipoCliente.Cliente
        Case vbKeyF11: CargoDatosCliente paClienteEmpresa
    End Select
End Sub

Private Sub TCI_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
    
        Dim aCi As String
        Screen.MousePointer = 11
        
        If Len(tCi.Text) = 7 Then tCi.Text = clsGeneral.AgregoDigitoControlCI(tCi.Text)
                
        'Valido la Cédula ingresada----------
        If Trim(tCi.Text) <> "" Then
            If Len(tCi.Text) <> 8 Then
                Screen.MousePointer = 0
                MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN": Exit Sub
            End If
            If Not clsGeneral.CedulaValida(tCi.Text) Then
                Screen.MousePointer = 0
                MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN": Exit Sub
            End If
        End If
        
        'Busco el Cliente -----------------------
        If Trim(tCi.Text) <> "" Then
            gCliente = BuscoClienteCIRUC(tCi.Text)
            If gCliente = 0 Then
                LimpioFichaCliente
                Screen.MousePointer = 0
                MsgBox "No existe un cliente para la cédula ingresada.", vbExclamation, "ATENCIÓN"
            Else
                 CargoDatosCliente gCliente
            End If
        Else
            tRuc.SetFocus
        End If
        Screen.MousePointer = 0
    End If

End Sub

Private Sub tCI1_Change()
    tCI1.Tag = "0"
End Sub

Private Sub tCi1_GotFocus()
    tCI1.SelStart = 0: tCI1.SelLength = 11
End Sub

Private Sub tCi1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 93: PopupMenu MnuBuscar, , tCI1.Left + (tCI1.Width / 2), (tCI1.Top + tCI1.Height) - (tCI1.Height / 2)
        Case vbKeyF3: If Shift = 0 Then NuevoCliente TipoCliente.Cliente
        Case vbKeyF4: If Shift = 0 Then BuscarClientes TipoCliente.Cliente
        Case vbKeyF11: CargoDatosCliente paClienteEmpresa
    End Select
End Sub

Private Sub TCI1_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
    
        Dim aCi As String
        Screen.MousePointer = 11
        
        If Len(tCI1.Text) = 7 Then tCI1.Text = clsGeneral.AgregoDigitoControlCI(tCI1.Text)
                
        'Valido la Cédula ingresada----------
        If Trim(tCI1.Text) <> "" Then
            If Len(tCI1.Text) <> 8 Then
                Screen.MousePointer = 0
                MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
            If Not clsGeneral.CedulaValida(tCI1.Text) Then
                Screen.MousePointer = 0
                MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
        End If
        
        'Busco el Cliente -----------------------
        If Trim(tCI1.Text) <> "" Then
            gCliente = BuscoClienteCIRUC(tCI1.Text)
            If gCliente = 0 Then
                LimpioFichaCliente1
                Screen.MousePointer = 0
                MsgBox "No existe un cliente para la cédula ingresada.", vbExclamation, "ATENCIÓN"
            Else
                 CargoDatosCliente gCliente
            End If
        Else
            tRUC1.SetFocus
        End If
        Screen.MousePointer = 0
    End If

End Sub

Private Sub tRuc_Change()
    tRuc.Tag = "0"
End Sub

Private Sub tRuc_GotFocus()
    tRuc.SelStart = 0: tRuc.SelLength = 15
End Sub

Private Sub tRuc_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case 93: PopupMenu MnuBuscar, , tRuc.Left + (tRuc.Width / 2), (tRuc.Top + tCi.Height) - (tRuc.Height / 2), MnuBusquedas
        Case vbKeyF3: If Shift = 0 Then NuevoCliente TipoCliente.Empresa
        Case vbKeyF4: If Shift = 0 Then BuscarClientes TipoCliente.Empresa
        Case vbKeyF11: CargoDatosCliente paClienteEmpresa
    End Select
    
End Sub

Private Sub tRuc_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tRuc.Tag) = Trim(tRuc.Text) Then tCi.SetFocus: Exit Sub
        
        If Trim(tRuc.Text) <> "" Then
            Screen.MousePointer = 11
            gCliente = BuscoClienteCIRUC(Trim(tRuc.Text))
            If gCliente = 0 Then
                LimpioFichaCliente
                Screen.MousePointer = 0
                MsgBox "No existe un cliente para el número de RUC ingresado.", vbExclamation, "ATENCIÓN"
            Else
                'Cargo Datos del Cliente Seleccionado------------------------------------------------
                 CargoDatosCliente gCliente
            End If
        Else
            tCi.SetFocus
        End If
        Screen.MousePointer = 0
    End If
    
End Sub

Private Sub tRUC1_Change()
    tRUC1.Tag = "0"
End Sub

Private Sub tRUC1_GotFocus()
    tRUC1.SelStart = 0: tRUC1.SelLength = 15
End Sub

Private Sub tRUC1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 93: PopupMenu MnuBuscar, , tRUC1.Left + (tRUC1.Width / 2), (tRUC1.Top + tCI1.Height) - (tRUC1.Height / 2), MnuBusquedas
        Case vbKeyF3: If Shift = 0 Then NuevoCliente TipoCliente.Empresa
        Case vbKeyF4: If Shift = 0 Then BuscarClientes TipoCliente.Empresa
        Case vbKeyF11: CargoDatosCliente paClienteEmpresa
    End Select
End Sub

Private Sub tRUC1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(tRUC1.Tag) = Trim(tRUC1.Text) Then tCI1.SetFocus: Exit Sub
        
        If Trim(tRUC1.Text) <> "" Then
            Screen.MousePointer = 11
            gCliente = BuscoClienteCIRUC(Trim(tRUC1.Text))
            If gCliente = 0 Then
                LimpioFichaCliente1
                Screen.MousePointer = 0
                MsgBox "No existe un cliente para el número de RUC ingresado.", vbExclamation, "ATENCIÓN"
            Else
                'Cargo Datos del Cliente Seleccionado------------------------------------------------
                 CargoDatosCliente gCliente
            End If
        Else
            tCI1.SetFocus
        End If
        Screen.MousePointer = 0
    End If
End Sub

Private Sub CargoDatosCliente(IdCliente As Long)

    Cons = "Select * from Cliente " _
                & " Left Outer Join CPersona ON CliCodigo = CPeCliente " _
                & " Left Outer Join CEmpresa ON CliCodigo = CEmCliente " _
           & " Where CliCodigo = " & IdCliente
           
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    If Not RsAux.EOF Then       'CI o RUC
        Select Case RsAux!CliTipo
            Case TipoCliente.Cliente
                If Screen.ActiveControl.Name = tCi.Name Or Screen.ActiveControl.Name = tRuc.Name Then
                    LimpioFichaCliente
                    If Val(tCI1.Tag) = RsAux!CliCodigo Then MsgBox "Selecciono el mismo cliente, verifique", vbInformation, "ATENCIÓN": RsAux.Close: Exit Sub
                    'Ficha de arriba
                    If Not IsNull(RsAux!CliCiRuc) Then tCi.Text = clsGeneral.RetornoFormatoCedula(RsAux!CliCiRuc) Else tCi.Text = ""
                    tCi.Tag = RsAux!CliCodigo
                    tRuc.Text = "": tRuc.Tag = ""
                    lTitular.Caption = Trim(Trim(Format(RsAux!CPeNombre1, "#")) & " " & Trim(Format(RsAux!CPeNombre2, "#"))) & ", " & Trim(Trim(Format(RsAux!CPeApellido1, "#")) & " " & Trim(Format(RsAux!CPeApellido2, "#")))
                Else
                    LimpioFichaCliente1
                    If Val(tCi.Tag) = RsAux!CliCodigo Then MsgBox "Selecciono el mismo cliente, verifique", vbInformation, "ATENCIÓN": RsAux.Close: Exit Sub
                    If Not IsNull(RsAux!CliCiRuc) Then tCI1.Text = clsGeneral.RetornoFormatoCedula(RsAux!CliCiRuc) Else tCI1.Text = ""
                    tCI1.Tag = RsAux!CliCodigo
                    tRUC1.Text = "": tRUC1.Tag = ""
                    lTitular1.Caption = Trim(Trim(Format(RsAux!CPeNombre1, "#")) & " " & Trim(Format(RsAux!CPeNombre2, "#"))) & ", " & Trim(Trim(Format(RsAux!CPeApellido1, "#")) & " " & Trim(Format(RsAux!CPeApellido2, "#")))
                End If
            Case TipoCliente.Empresa
                If Screen.ActiveControl.Name = tRuc.Name Or Screen.ActiveControl.Name = tCi.Name Then
                    LimpioFichaCliente
                    If Val(tRUC1.Tag) = RsAux!CliCodigo Then MsgBox "Selecciono el mismo cliente, verifique", vbInformation, "ATENCIÓN": RsAux.Close: Exit Sub
                    If Not IsNull(RsAux!CliCiRuc) Then tRuc.Text = Trim(RsAux!CliCiRuc)
                    tRuc.Tag = RsAux!CliCodigo
                    tCi.Text = "": tCi.Tag = ""
                    If Not IsNull(RsAux!CEmNombre) Then lTitular.Caption = Trim(RsAux!CEmFantasia)
                    If Not IsNull(RsAux!CEmFantasia) Then lTitular.Caption = lTitular.Caption & " (" & Trim(RsAux!CEmFantasia) & ")"
                Else
                    LimpioFichaCliente1
                    If Val(tRuc.Tag) = RsAux!CliCodigo Then MsgBox "Selecciono el mismo cliente, verifique", vbInformation, "ATENCIÓN": RsAux.Close: Exit Sub
                    If Not IsNull(RsAux!CliCiRuc) Then tRUC1.Text = Trim(RsAux!CliCiRuc)
                    tRUC1.Tag = RsAux!CliCodigo
                    tCI1.Text = "": tCI1.Tag = ""
                    If Not IsNull(RsAux!CEmNombre) Then lTitular1.Caption = Trim(RsAux!CEmFantasia)
                    If Not IsNull(RsAux!CEmFantasia) Then lTitular1.Caption = lTitular1.Caption & " (" & Trim(RsAux!CEmFantasia) & ")"
                End If
        End Select
        'Direccion
        If Not IsNull(RsAux!CliDireccion) Then
            If Screen.ActiveControl.Name = tCi.Name Or Screen.ActiveControl.Name = tRuc.Name Then
                lDireccion.Caption = clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!CliDireccion, Departamento:=True, Localidad:=True, Zona:=True, ConfYVD:=True)
                lTelefono.Caption = TelefonoATexto(IdCliente)     'Telefonos
            Else
                lDireccion1.Caption = clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!CliDireccion, Departamento:=True, Localidad:=True, Zona:=True, ConfYVD:=True)
                lTelefono1.Caption = TelefonoATexto(IdCliente)     'Telefonos
            End If
        End If
        gCliente = RsAux!CliCodigo
        RsAux.Close
        'Verifico que no sea el mismo cliente.
        If Screen.ActiveControl.Name = tCi.Name Or Screen.ActiveControl.Name = tRuc.Name Then
            CargoProducto gCliente, vsProducto
            Foco tCI1
        Else
            CargoProducto gCliente, vsProducto1
        End If
    Else
        RsAux.Close
        If Screen.ActiveControl.Name = tCi.Name Or Screen.ActiveControl.Name = tRuc.Name Then
            tRuc.Text = "": tRuc.Tag = ""
            tCi.Text = "": tCi.Tag = ""
        Else
            tRUC1.Text = "": tRUC1.Tag = ""
            tCI1.Text = "": tCI1.Tag = ""
        End If
    End If
    Exit Sub
errCliente:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del cliente."
End Sub

Private Sub BuscarClientes(aTipoCliente As Integer)
    
    Screen.MousePointer = 11
    Dim objBuscar As New clsBuscarCliente
    Dim aTipo As Integer, aCliente As Long
    
    If aTipoCliente = TipoCliente.Cliente Then objBuscar.ActivoFormularioBuscarClientes cBase, Persona:=True
    If aTipoCliente = TipoCliente.Empresa Then objBuscar.ActivoFormularioBuscarClientes cBase, Empresa:=True
    Me.Refresh
    aTipo = objBuscar.BCTipoClienteSeleccionado
    aCliente = objBuscar.BCClienteSeleccionado
    Set objBuscar = Nothing
    
    On Error GoTo errCargar
    If aCliente <> 0 Then CargoDatosCliente aCliente
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub NuevoCliente(aTipoCliente As Integer)
On Error GoTo ErrFC
    Dim aCliente As Long
    Screen.MousePointer = 11
    Dim objCliente As New clsCliente
    If aTipoCliente = TipoCliente.Cliente Then
        objCliente.Personas 0, 0, 1
    Else
        objCliente.Empresas 0, True
    End If
    Me.Refresh
    aCliente = objCliente.IDIngresado
    Set objCliente = Nothing
    CargoDatosCliente aCliente
    Screen.MousePointer = 0
    Exit Sub
ErrFC:
    clsGeneral.OcurrioError "Ocurrio un error al ir a ficha de cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function BuscoClienteCIRUC(CiRuc As String)

    On Error GoTo errBuscar
    BuscoClienteCIRUC = 0
    Cons = "Select * from Cliente Where CliCiRuc = '" & Trim(CiRuc) & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then BuscoClienteCIRUC = RsAux!CliCodigo
    RsAux.Close
    Exit Function

errBuscar:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el cliente."
    Screen.MousePointer = 0
End Function

Private Sub LimpioFichaCliente()
    'Datos del cliente.----------------------------
    lTitular.Caption = ""
    lDireccion.Caption = ""
    lTelefono.Caption = ""
    tRuc.Tag = ""
    tCi.Tag = ""
    vsProducto.Rows = 1
End Sub

Private Sub LimpioFichaCliente1()
    'Datos del cliente.----------------------------
    lTitular1.Caption = ""
    lDireccion1.Caption = ""
    lTelefono1.Caption = ""
    tRUC1.Tag = ""
    tCI1.Tag = ""
    vsProducto1.Rows = 1
End Sub

Private Sub CargoProducto(IdCliente As Long, vsGrilla As vsFlexGrid)
On Error GoTo ErrCDP
Dim aValor As Integer, fModificado As String, aCantP As Long
Dim aCodNomArt As String, nroIntentos As Integer
Dim objLista As clsListadeAyuda

    Screen.MousePointer = 11
    vsGrilla.Rows = 1
    
    nroIntentos = 0
    
    Cons = "Select Count(*) From Producto Where ProCliente = " & IdCliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not IsNull(RsAux(0)) Then aCantP = RsAux(0)
    RsAux.Close
    
    'Veo si es el cliente que va a ceder o el que va a recibir.
    'Si es el primero y tiene + de 10 Productos le pido el artículo a cargar
    'Si es el sgdo. y posee + de 10 Productos no cargo la lista.
    
    If LCase(vsGrilla.Name) = "vsproducto" Then
        
        If aCantP > 10 Then
            Do While aCodNomArt = ""
                nroIntentos = nroIntentos + 1
                aCodNomArt = InputBox("Ingrese el código o el nombre del artículo que desea ceder.", "Cliente con varios artículos")
                If aCodNomArt <> "" Then
                    If IsNumeric(aCodNomArt) Then
                        'Busco por código
                        Cons = "Select * From Articulo Where ArtCodigo = " & Val(aCodNomArt)
                        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                        If RsAux.EOF Then
                            RsAux.Close
                            MsgBox "No existe un artículo con el código ingresado.", vbInformation, "ATENCIÓN"
                            aCodNomArt = ""
                        Else
                            aCodNomArt = RsAux!ArtID
                            RsAux.Close
                        End If
                    Else
                        'Busco por nombre
                        Cons = "Select ArtID, ArtCodigo as 'Código', ArtNombre as 'Nombre' From Articulo Where ArtNombre Like '" & aCodNomArt & "%'"
                        Set objLista = New clsListadeAyuda
                        If objLista.ActivarAyuda(cBase, Cons, 400, 1) > 0 Then
                            aCodNomArt = objLista.RetornoDatoSeleccionado(0)
                        End If
                        Set objLista = Nothing
                        If Val(aCodNomArt) = 0 Then aCodNomArt = ""
                    End If
                End If
                If aCodNomArt = "" And nroIntentos = 3 Then
                    Select Case MsgBox("Presione: " & vbCr & vbTab _
                            & "<Si> Para seguir intentando buscar el artículo." & vbCr & vbTab _
                            & "<NO> Para cargar todos los artículos del cliente." & vbCr & vbTab _
                            & "<Cancelar> Para cargar solamente los artículos del cliente modificados hoy.", vbQuestion + vbYesNoCancel, "VARIOS INTENTOS")
                        
                        Case vbYes: nroIntentos = 0
                        Case vbNo: Exit Do
                        Case vbCancel: aCodNomArt = "-2"
                    End Select
                End If
            Loop
        End If
        
    Else
        'Es el sgdo.
        If aCantP > 10 Then aCodNomArt = "-2"
    End If
    
    Cons = "Select * From Producto, Articulo " _
        & " Where ProCliente = " & IdCliente & " And ProArticulo = ArtID"
    
    If Val(aCodNomArt) <> 0 Then
        If Val(aCodNomArt) = -2 Then
            Cons = Cons & " And ProFModificacion > '" & Format(Date - 1, "mm/dd/yyyy 23:59:59") & "'"
        Else
            Cons = Cons & " And ArtID = " & Val(aCodNomArt)
        End If
    End If
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    Do While Not RsAux.EOF
        With vsGrilla
            .AddItem ""
            
            aValor = RsAux!ArtID
            .Cell(flexcpData, .Rows - 1, 0) = aValor
            fModificado = RsAux!ProFModificacion
            .Cell(flexcpData, .Rows - 1, 1) = fModificado
            
            aValor = CalculoEstadoProducto(RsAux!ProCodigo)
            .Cell(flexcpData, .Rows - 1, 3) = aValor    'Estado del producto para cargar combo automático.
            
            .Cell(flexcpText, .Rows - 1, 2) = EstadoProducto(CInt(aValor), True)
            
            .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!ProCodigo, "#,000")
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!ArtNombre)
            
            
            If Not IsNull(RsAux!ProCompra) Then .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!ProCompra, "dd/mm/yyyy")
            'Saco garantia-----------------------------------------------------------------------------------------------------------------------
            .Cell(flexcpText, .Rows - 1, 4) = RetornoGarantia(RsAux!ArtID)
            '--------------------------------------------------------------------------------------------------------------------------------------
            If Not IsNull(RsAux!ProNroSerie) Then .Cell(flexcpText, .Rows - 1, 5) = RsAux!ProNroSerie
            If Not IsNull(RsAux!ProFacturaS) Then .Cell(flexcpText, .Rows - 1, 6) = Trim(RsAux!ProFacturaS) & " "
            If Not IsNull(RsAux!ProFacturaN) Then .Cell(flexcpText, .Rows - 1, 6) = .Cell(flexcpText, .Rows - 1, 6) & Trim(RsAux!ProFacturaN)
            
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub
ErrCDP:
    clsGeneral.OcurrioError "Ocurrio un error al cargar los datos del producto.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Sub vsProducto_DblClick()
    If vsProducto.Row > 0 Then
        If Val(tCI1.Tag) > 0 Then CedoProducto Val(tCI1.Tag), vsProducto, 0: Exit Sub
        If Val(tRUC1.Tag) > 0 Then CedoProducto Val(tRUC1.Tag), vsProducto, 0: Exit Sub
    End If
End Sub

Private Sub vsProducto_GotFocus()
    Status.SimpleText = "Seleccione un producto de la lista y con doble click o enter lo cede al otro cliente."
End Sub

Private Sub vsProducto_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsProducto.Rows > 1 Then
        If KeyCode = vbKeyReturn Then
            If Val(tCI1.Tag) > 0 Then CedoProducto Val(tCI1.Tag), vsProducto, 1: Exit Sub
            If Val(tRUC1.Tag) > 0 Then CedoProducto Val(tRUC1.Tag), vsProducto, 1: Exit Sub
        End If
    End If
End Sub

Private Sub vsProducto1_DblClick()
    If vsProducto1.Row > 0 Then
        If Val(tCi.Tag) > 0 Then CedoProducto Val(tCi.Tag), vsProducto1, 1: Exit Sub
        If Val(tRuc.Tag) > 0 Then CedoProducto Val(tRuc.Tag), vsProducto1, 1: Exit Sub
    End If
End Sub

Private Sub vsProducto1_GotFocus()
    Status.SimpleText = "Seleccione un producto de la lista y con doble click o enter lo cede al otro cliente."
End Sub

Private Sub vsProducto1_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsProducto1.Rows > 1 Then
        If KeyCode = vbKeyReturn Then
            If Val(tCi.Tag) > 0 Then CedoProducto Val(tCi.Tag), vsProducto1, 1: Exit Sub
            If Val(tRuc.Tag) > 0 Then CedoProducto Val(tRuc.Tag), vsProducto1, 1: Exit Sub
        End If
    End If
End Sub

Private Sub CedoProducto(IdCliente As Long, vsGrilla As vsFlexGrid, Caso As Integer)
On Error GoTo ErrGrabo
Dim UsuID As Long, sDefensa As String, sTexto As String
Dim idCliQueDa As Long

    If MsgBox("¿Confirma ceder el producto seleccionado?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
        
        UsuID = 0: sDefensa = ""
        GeneroSuceso UsuID, sDefensa
        
        If UsuID = 0 Then
            Screen.MousePointer = 0
            MsgBox "Debe ingresar el suceso.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
            
        Screen.MousePointer = 11
        
        Cons = "Select * From Producto Where ProCodigo = " & vsGrilla.Cell(flexcpValue, vsGrilla.Row, 0)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux.EOF Then
            RsAux.Close
            MsgBox "No se encontro el producto seleccionado.", vbExclamation, "ATENCIÓN"
        Else
            If RsAux!ProFModificacion = CDate(vsGrilla.Cell(flexcpData, vsGrilla.Row, 1)) Then
                
                RsAux.Edit
                RsAux!ProCliente = IdCliente
                RsAux!ProFModificacion = Format(gFechaServidor, sqlFormatoFH)
                RsAux.Update
                RsAux.Close
                
                If Caso = 0 Then
                    If Val(tCi.Tag) > 0 Then
                        idCliQueDa = Val(tCi.Tag)
                    Else
                        idCliQueDa = Val(tRuc.Tag)
                    End If
                Else
                    If Val(tCI1.Tag) > 0 Then
                        idCliQueDa = Val(tCI1.Tag)
                    Else
                        idCliQueDa = Val(tRUC1.Tag)
                    End If
                End If
                
                sTexto = "Da cliente id: " & idCliQueDa & " recibe cliente id: " & IdCliente
                
                'Registro el suceso
                clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.CederProductoServicio, paCodigoDeTerminal, UsuID, 0, vsGrilla.Cell(flexcpData, vsGrilla.Row - 1, 0), sTexto, Trim(sDefensa), 1
                
                
            Else
                RsAux.Close
                MsgBox "Otra terminal modificó los datos del producto, verifique los cambios.", vbExclamation, "ATENCIÓN"
            End If
        End If
        If Val(tCi.Tag) > 0 Then gCliente = Val(tCi.Tag) Else gCliente = Val(tRuc.Tag)
        CargoProducto gCliente, vsProducto
        If Val(tCI1.Tag) > 0 Then gCliente = Val(tCI1.Tag) Else gCliente = Val(tRUC1.Tag)
        CargoProducto gCliente, vsProducto1
        Screen.MousePointer = 0
    End If
    Exit Sub
ErrGrabo:
    clsGeneral.OcurrioError "Ocurrio un error al intentar grabar.", Err.Description
End Sub
Private Sub CargoDatosClienteDado(IdCliente As Long)

    Cons = "Select * from Cliente " _
                & " Left Outer Join CPersona ON CliCodigo = CPeCliente " _
                & " Left Outer Join CEmpresa ON CliCodigo = CEmCliente " _
           & " Where CliCodigo = " & IdCliente
           
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    If Not RsAux.EOF Then       'CI o RUC
        Select Case RsAux!CliTipo
            Case TipoCliente.Cliente
                    LimpioFichaCliente
                    'Ficha de arriba
                    If Not IsNull(RsAux!CliCiRuc) Then tCi.Text = clsGeneral.RetornoFormatoCedula(RsAux!CliCiRuc) Else tCi.Text = ""
                    tCi.Tag = RsAux!CliCodigo
                    tRuc.Text = "": tRuc.Tag = ""
                    lTitular.Caption = Trim(Trim(Format(RsAux!CPeNombre1, "#")) & " " & Trim(Format(RsAux!CPeNombre2, "#"))) & ", " & Trim(Trim(Format(RsAux!CPeApellido1, "#")) & " " & Trim(Format(RsAux!CPeApellido2, "#")))
            Case TipoCliente.Empresa
                    LimpioFichaCliente
                    If Not IsNull(RsAux!CliCiRuc) Then tRuc.Text = Trim(RsAux!CliCiRuc)
                    tRuc.Tag = RsAux!CliCodigo
                    tCi.Text = "": tCi.Tag = ""
                    If Not IsNull(RsAux!CEmNombre) Then lTitular.Caption = Trim(RsAux!CEmFantasia)
                    If Not IsNull(RsAux!CEmFantasia) Then lTitular.Caption = lTitular.Caption & " (" & Trim(RsAux!CEmFantasia) & ")"
        End Select
        'Direccion
        If Not IsNull(RsAux!CliDireccion) Then
            lDireccion.Caption = clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!CliDireccion, Departamento:=True, Localidad:=True, Zona:=True, ConfYVD:=True)
            lTelefono.Caption = TelefonoATexto(IdCliente)     'Telefonos
        End If
        gCliente = RsAux!CliCodigo
        RsAux.Close
        'Verifico que no sea el mismo cliente.
        CargoProducto gCliente, vsProducto
    Else
        RsAux.Close
    End If
    Exit Sub
errCliente:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del cliente.", Err.Description
End Sub

Private Sub GeneroSuceso(idUsuario As Long, sDefensa As String)
        
    Dim objSuceso As New clsSuceso
    idUsuario = 0: sDefensa = ""
    objSuceso.ActivoFormulario paCodigoDeUsuario, "Ceder producto", cBase
    Me.Refresh
    idUsuario = objSuceso.RetornoValor(True)
    sDefensa = objSuceso.RetornoValor(False, True)
    Set objSuceso = Nothing
    
End Sub

