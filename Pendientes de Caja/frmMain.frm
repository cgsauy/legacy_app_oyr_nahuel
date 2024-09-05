VERSION 5.00
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D9D9E0F6-C86B-4B3A-BFD9-06B9B5B7A222}#2.1#0"; "orUserDigit.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pendientes de Caja"
   ClientHeight    =   5970
   ClientLeft      =   3645
   ClientTop       =   2865
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7200
   Begin VB.TextBox tLClave 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   305
      IMEMode         =   3  'DISABLE
      Left            =   5700
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   28
      Top             =   5340
      Width           =   1365
   End
   Begin VB.TextBox tPClave 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   305
      IMEMode         =   3  'DISABLE
      Left            =   2820
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   19
      Top             =   3720
      Width           =   1365
   End
   Begin VB.TextBox tEnvio 
      Height          =   300
      Left            =   1020
      MaxLength       =   8
      TabIndex        =   14
      Top             =   2280
      Width           =   1035
   End
   Begin VB.TextBox tLFecha 
      Height          =   300
      Left            =   1080
      TabIndex        =   25
      Top             =   5340
      Width           =   1155
   End
   Begin VB.TextBox tLMemo 
      Height          =   660
      Left            =   120
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   4620
      Width           =   6975
   End
   Begin MSComCtl2.DTPicker dOrigen 
      Height          =   315
      Left            =   1020
      TabIndex        =   5
      Top             =   1140
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   80478209
      CurrentDate     =   37543
   End
   Begin VB.TextBox tImporte 
      Height          =   300
      Left            =   5460
      MaxLength       =   12
      TabIndex        =   9
      Text            =   "1.999.999.00"
      Top             =   1500
      Width           =   1575
   End
   Begin VB.TextBox tID 
      Height          =   300
      Left            =   1020
      MaxLength       =   7
      TabIndex        =   1
      Top             =   480
      Width           =   1035
   End
   Begin orUserDigit.UserDigit tLUsuario 
      Height          =   285
      Left            =   3660
      TabIndex        =   27
      Top             =   5340
      Width           =   1935
      _ExtentX        =   2805
      _ExtentY        =   503
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
   Begin VB.TextBox tSerie 
      Height          =   300
      Left            =   1020
      MaxLength       =   2
      TabIndex        =   11
      Text            =   "S"
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox tNumero 
      Height          =   300
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   12
      Top             =   1920
      Width           =   1035
   End
   Begin MSComctlLib.StatusBar sbHelp 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   29
      Top             =   5700
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12197
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox tPMemo 
      Height          =   660
      Left            =   120
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   2940
      Width           =   6975
   End
   Begin AACombo99.AACombo cDisponibilidad 
      Height          =   315
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Width           =   3555
      _ExtentX        =   6271
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6000
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D86
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin AACombo99.AACombo cConcepto 
      Height          =   315
      Left            =   1020
      TabIndex        =   7
      Top             =   1500
      Width           =   3315
      _ExtentX        =   5847
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
   Begin orUserDigit.UserDigit tPAutoriza 
      Height          =   285
      Left            =   5160
      TabIndex        =   21
      Top             =   3720
      Width           =   1935
      _ExtentX        =   2805
      _ExtentY        =   503
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
   Begin orUserDigit.UserDigit tPUsuario 
      Height          =   285
      Left            =   840
      TabIndex        =   18
      Top             =   3720
      Width           =   1935
      _ExtentX        =   2805
      _ExtentY        =   503
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº &Envío:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2340
      Width           =   795
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ingresado el "
      Height          =   255
      Left            =   4380
      TabIndex        =   37
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lFIngreso 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5460
      TabIndex        =   36
      Top             =   1140
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "&Comentarios"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4380
      Width           =   1155
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3780
      Width           =   675
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Autoriza:"
      Height          =   255
      Left            =   4440
      TabIndex        =   20
      Top             =   3780
      Width           =   675
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "&Importe:"
      Height          =   255
      Left            =   4740
      TabIndex        =   8
      Top             =   1560
      Width           =   795
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Conce&pto:"
      Height          =   255
      Left            =   -120
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "&ID Pend.:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   1035
   End
   Begin VB.Label lLLiquidacion 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   15
      Left            =   2580
      TabIndex        =   35
      Top             =   4185
      Width           =   4755
   End
   Begin VB.Label lTLiquidacion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos de Liquidación"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   34
      Top             =   4080
      Width           =   1470
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario:"
      Height          =   255
      Left            =   2940
      TabIndex        =   26
      Top             =   5400
      Width           =   675
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "&Disponibilidad:"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   540
      Width           =   1035
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Liquidado el:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label lFactura 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   2520
      TabIndex        =   32
      Top             =   1920
      Width           =   4515
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Docu&mento:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1980
      Width           =   915
   End
   Begin VB.Label lLPendiente 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   15
      Left            =   2640
      TabIndex        =   31
      Top             =   1005
      Width           =   4755
   End
   Begin VB.Label lTPendiente 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos del Pendiente"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   900
      Width           =   1440
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "&Comentarios"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2700
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "F/&Origen:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   675
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
         Visible         =   0   'False
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
   Begin VB.Menu MnuAccesos 
      Caption         =   "&Accesos"
      Begin VB.Menu MnuDetalle 
         Caption         =   "Detalle de Factura"
      End
      Begin VB.Menu MnuVisualizacion 
         Caption         =   "Visualización de Operaciones"
      End
      Begin VB.Menu MnuAccL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPlantillas 
         Caption         =   "Ver Todos los Pendientes"
      End
   End
   Begin VB.Menu MnuExit 
      Caption         =   "&Salir"
      Begin VB.Menu MnuSalir 
         Caption         =   "Del formulario"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuAyuda 
      Caption         =   "&?"
      Begin VB.Menu MnuHelp 
         Caption         =   "&Ayuda"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim prmNuevo As Boolean, prmLiquidar As Boolean
Dim mTexto As String

Public Function gbl_CargaConParametros(XID As Long, XidDoc As Long)

    If XID = 0 And XidDoc = 0 Then Exit Function
    
    If XID <> 0 Then
        CargoPendiente XID
        Exit Function
    End If
    
    If XidDoc <> 0 Then AccionNuevo mDocumento:=XidDoc
    
End Function

Private Sub cConcepto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cConcepto.ListIndex <> -1 Then Foco tImporte
End Sub

Private Sub cDisponibilidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cDisponibilidad.ListIndex <> -1 Then
            If dOrigen.Enabled Then dOrigen.SetFocus
        End If
    End If
End Sub

Private Sub dOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        On Error Resume Next
        
        If Val(lFactura.Tag) <> 0 Then
            If Trim(tSerie.Tag) <> Format(dOrigen.Value, "dd/mm/yyyy") Then
                tSerie.Text = "": tNumero.Text = ""
            End If
        End If
                
        cConcepto.SetFocus
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    ObtengoSeteoForm Me
    InicializoForm
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    lLPendiente.Left = lTPendiente.Left + lTPendiente.Width + 80
    lLPendiente.Width = Me.ScaleWidth - lLPendiente.Left - 80
    
    lLLiquidacion.Left = lTLiquidacion.Left + lTLiquidacion.Width + 80
    lLLiquidacion.Width = Me.ScaleWidth - lLLiquidacion.Left - 80
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    GuardoSeteoForm Me
    EndMain
End Sub

Private Sub InicializoForm()
    
    On Error Resume Next
    FechaDelServidor
    LimpioFicha
    
    'Cargo las Disponibilidades-------------------------------------------------------------------
    Dim mDisME As String
    cons = "Select * from Sucursal Where SucDisponibilidadME Is not Null"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        If Trim(rsAux!SucDisponibilidadME) <> "" Then
            If mDisME <> "" Then
                If Right(mDisME, 1) <> "," Then mDisME = mDisME & ","
            End If
            mDisME = mDisME & Trim(rsAux!SucDisponibilidadME)
        End If
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    cons = "Select DisID, DisNombre from Disponibilidad " _
            & " Where DisID In (Select SucDisponibilidad from Sucursal) "
    If Trim(mDisME) <> "" Then cons = cons & " OR DisID In (" & mDisME & ") "
    cons = cons & " Order by DisNombre"
    
    CargoCombo cons, cDisponibilidad, ""
    '-----------------------------------------------------------------------------------------
    
    
    cons = "Select TPeCodigo, TPeNombre from TipoPendiente Order by TPeNombre"
    CargoCombo cons, cConcepto
    
    With tPUsuario
        .Connect cBase
        .UserID = 0: .EnabledButton = True
        .Terminal = paCodigoDeTerminal
        .UserLog = paCodigoDeUsuario
        .SetApp "Pendientes Caja"
        .ControlKey = .Name
        '.EnabledButton = False
    End With
    With tPAutoriza
        .Connect cBase
        .UserID = 0: .EnabledButton = False
    End With
    With tLUsuario
        .Connect cBase
        .UserID = 0: .EnabledButton = True
        .Terminal = paCodigoDeTerminal
        .UserLog = paCodigoDeUsuario
        .SetApp "Pendientes Caja"
        .ControlKey = .Name
        '.EnabledButton = False
    End With
    
    HabilitoCampos
    
End Sub

Private Sub LimpioFicha()
    
    prmNuevo = False: prmLiquidar = False
    
    lFIngreso.Caption = ""
    cDisponibilidad.Text = ""
    dOrigen.Value = Format(gFechaServidor, "dd/mm/yyyy")
    cConcepto.Text = ""
    tImporte.Text = ""
    tSerie.Text = "": tNumero.Text = ""
    tEnvio.Text = ""
    tPMemo.Text = ""
    
    tPUsuario.UserName = "": tPUsuario.UserID = 0
    tPClave.Text = ""
    tPAutoriza.UserName = "": tPAutoriza.UserID = 0
    
    tLMemo.Text = ""
    tLFecha.Text = ""
    tLUsuario.UserName = "": tLUsuario.UserID = 0
    tLClave.Text = ""
    
End Sub


Private Sub Label14_Click()
    Foco cDisponibilidad
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuDetalle_Click()
    EjecutarApp prmPathApp & "\Detalle de Factura.exe", IIf(Val(lFactura.Tag) <> 0, lFactura.Tag, "")
End Sub

Private Sub MnuEliminar_Click()
    AccionEliminar
End Sub

Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuHelp_Click()
    AccionMenuHelp
End Sub

Private Sub MnuNuevo_Click()
    AccionNuevo
End Sub

Private Sub MnuPlantillas_Click()
    EjecutarApp prmPathApp & "\appExploreMsg.exe ", prmPlantillas & ":0"
End Sub

Private Sub MnuSalir_Click()
    Unload Me
End Sub

Private Function ValidoGrabar() As Boolean

    ValidoGrabar = False
    
    If cDisponibilidad.ListIndex = -1 Then
        MsgBox "Falta seleccionar la disponibilidad para ingresar el pendiente.", vbExclamation, "Falta Disponibilidad"
        cDisponibilidad.SetFocus: Exit Function
    End If
    
    If Not IsDate(dOrigen.Value) Then
        MsgBox "Falta ingresar la fecha que se originó pendiente.", vbExclamation, "Falta Fecha de Pendiente"
        dOrigen.SetFocus: Exit Function
    End If
    
    If Format(dOrigen.Value, "yyyymmdd") > Format(gFechaServidor, "yyyymmdd") Then
        MsgBox "La fecha que se originó pendiente no puede ser mayor al día de hoy.", vbExclamation, "Fecha de Pendiente"
        dOrigen.SetFocus: Exit Function
    End If
        
    If prmNuevo Then
        'Valido que la disponibilidad no este cerrada a la fecha del pendiente
        If fnc_FechaCierreDisp(cDisponibilidad.ItemData(cDisponibilidad.ListIndex)) >= dOrigen.Value Then
            MsgBox "La fecha del pendiente es incorrecta, la disponibilidad ya fue cerrada.", vbExclamation, "Error en Fecha de Pendiente"
            dOrigen.SetFocus: Exit Function
        End If
    
'        If Val(lFactura.Tag) <> 0 Then
'            If Trim(tSerie.Tag) <> Format(dOrigen.Value, "dd/mm/yyyy") Then
'                MsgBox "Ud ingresó el pendiente para un documento." & vbCrLf & _
                            "La fecha del pendiente debe ser igual a la del documento. Verifique.", vbExclamation, "Error en Fecha de Pendiente"
'                dOrigen.SetFocus: Exit Function
'            End If
'        End If
        
        If Trim(tEnvio.Text) <> "" Then
            If Not IsNumeric(tEnvio.Text) Then
                MsgBox "El número de envío ingresado no es correcto.", vbExclamation, "Posible Error"
                Foco tEnvio: Exit Function
            End If
        End If
            
        'If Not miConexion.ValidoClave(Val(tPUsuario.UserID), Trim(UCase(tPClave.Text))) Then
        If Not miConexion.ValidoClave(Val(tPUsuario.UserID), Trim(tPClave.Text)) Then
            MsgBox "La clave de usuario ingresada no es correcta.", vbExclamation, "Error en Clave"
            Foco tPClave: Exit Function
        End If
    
    End If
    
    If cConcepto.ListIndex = -1 Then
        MsgBox "Falta seleccionar el concepto del pendiente.", vbExclamation, "Falta Concepto"
        cConcepto.SetFocus: Exit Function
    End If
       
    If Trim(tImporte.Text) = "" Or Not IsNumeric(tImporte.Text) Then
        MsgBox "El importe del pendiente no es correcto. Verifique.", vbExclamation, "Falta Importe del Pendiente"
        Foco tImporte: Exit Function
    End If
    
    If tPUsuario.UserID = 0 Then
        MsgBox "Falta ingresar el usuario que ingresa el pendiente.", vbExclamation, "Falta Usuario del Pendiente"
        tPUsuario.SetFocus: Exit Function
    End If
    
    If tPAutoriza.UserID = 0 Then
        MsgBox "Falta ingresar el usuario que autoriza el pendiente.", vbExclamation, "Falta Usuario Autoriza"
        tPAutoriza.SetFocus: Exit Function
    End If
    
    If tLUsuario.UserID = 0 And tLUsuario.Enabled Then
        MsgBox "Falta ingresar el usuario que liquida el pendiente.", vbExclamation, "Falta Usuario que Liquida"
        tLUsuario.SetFocus: Exit Function
    End If
    If tLUsuario.Enabled Then
        'If Not miConexion.ValidoClave(Val(tLUsuario.UserID), Trim(UCase(tLClave.Text))) Then
        If Not miConexion.ValidoClave(Val(tLUsuario.UserID), Trim(tLClave.Text)) Then
            MsgBox "La clave de usuario ingresada no es correcta.", vbExclamation, "Error en Clave"
            Foco tLClave: Exit Function
        End If
    End If
    ValidoGrabar = True
                
End Function

Private Sub AccionGrabar()

On Error GoTo errValidar
       
    sbHelp.Panels("help").Text = "Validando datos, espere ...": sbHelp.Refresh
    If Not ValidoGrabar Then
        sbHelp.Panels("help").Text = "": sbHelp.Refresh
        Exit Sub
    End If
    sbHelp.Panels("help").Text = "": sbHelp.Refresh
    
    If prmNuevo Then
        If MsgBox("Confirma grabar los datos del pendiente ?", vbQuestion + vbYesNo, "Grabar Datos") = vbNo Then Exit Sub
    End If
    
    If prmLiquidar Then
        If MsgBox("Confirma liquidar el Pendiente de Caja ?", vbQuestion + vbYesNo, "Liquidar Pendiente") = vbNo Then Exit Sub
    End If
    
    On Error GoTo errorBT
    
    FechaDelServidor

    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
    On Error GoTo errorET
    Screen.MousePointer = 11
    
    Dim mID As Long
    If Trim(tID.Text) <> "" Then mID = Trim(tID.Text) Else mID = 0
    
    cons = "Select * from PendientesCaja Where PCaID = " & mID
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If rsAux.EOF Then rsAux.AddNew Else rsAux.Edit
    
    If prmNuevo Then
        'Datos del Ingreso del Pendiente  --------------------------------------------------------------
        rsAux!PCaDisponibilidad = cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
        rsAux!PCaConcepto = cConcepto.ItemData(cConcepto.ListIndex)
        
        rsAux!PCaFPendiente = Format(dOrigen.Value, "mm/dd/yyyy")
        rsAux!PCaImporte = CCur(tImporte.Text)
        If Val(lFactura.Tag) <> 0 Then rsAux!PCaDocumento = Val(lFactura.Tag) Else rsAux!PCaDocumento = Null
        
        If Trim(tPMemo.Text) <> "" Then rsAux!PCaDetalleOrigen = Trim(tPMemo.Text) Else rsAux!PCaDetalleOrigen = Null
        
        rsAux!PCaUsrPendiente = tPUsuario.UserID
        rsAux!PCaUsrAutoriza = tPAutoriza.UserID
        rsAux!PCaFIngreso = Format(gFechaServidor, "mm/dd/yyyy hh:mm")
        
        If IsNumeric(tEnvio.Text) Then rsAux!PCaEnvio = tEnvio.Text Else rsAux!PCaEnvio = Null
    End If
    
    If prmLiquidar Then
        'Datos de la Liquidacion  --------------------------------------------------------------------------
        rsAux!PCaFLiquidacion = Format(tLFecha.Text, "mm/dd/yyyy")
        rsAux!PCaUsrLiquidacion = tLUsuario.UserID
        If Trim(tLMemo.Text) <> "" Then rsAux!PCaDetalleLiquidacion = Trim(tLMemo.Text) Else rsAux!PCaDetalleLiquidacion = Null
    End If
    rsAux.Update: rsAux.Close
    
    If mID = 0 Then
        'Saco id Pendiente  ------------------------------------------------------------------
        cons = "Select * from PendientesCaja " & _
                    " Where PCaDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex) & _
                    " And PCaConcepto = " & cConcepto.ItemData(cConcepto.ListIndex) & _
                    " And PCaFIngreso = " & Format(gFechaServidor, "'mm/dd/yyyy hh:mm'") & _
                    " And PCaUsrPendiente = " & tPUsuario.UserID
                    
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then mID = rsAux!PCaID
        rsAux.Close
    End If
    
    cBase.CommitTrans    'Fin de la TRANSACCION------------------------------------------
    
    If Trim(tID.Text) = "" And mID <> 0 Then tID.Text = mID
    AccionCancelar
    Screen.MousePointer = 0
    Exit Sub

errValidar:
    clsGeneral.OcurrioError "Error al procesar los datos para grabar.", Err.Description
    Screen.MousePointer = 0: Exit Sub
    
errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente.", Err.Description
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Error al grabar los datos del cambio.", Err.Description
    Screen.MousePointer = 0: Exit Sub
End Sub

Private Sub MnuVisualizacion_Click()
    EjecutarApp prmPathApp & "\Visualizacion de Operaciones.exe"
End Sub

Private Sub tEnvio_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tPMemo
End Sub

Private Sub tID_Change()
    
    If tID.Enabled Then
        If Val(tID.Tag) <> 0 Then
            Botones True, False, False, False, False, Toolbar1, Me
            tID.Tag = 0
        End If
    End If
    
End Sub

Private Sub tID_GotFocus()
    tID.SelStart = 0: tID.SelLength = Len(tID.Text)
    sbHelp.Panels("help").Text = "[F1]- Ver últimos pendientes.        [F2]- Ver pendientes liquidados.": sbHelp.Refresh
End Sub

Private Sub tID_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyF1: AyudaPendientes
        Case vbKeyF2: AyudaPendientes bLiquidados:=True
    End Select
    
End Sub

Private Sub tID_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tID.Text) Then
            If Trim(tID.Tag) <> Trim(tID.Text) Then CargoPendiente tID.Text
        End If
    End If
    
End Sub

Private Sub tID_LostFocus()
    sbHelp.Panels("help").Text = ""
End Sub

Private Sub tImporte_GotFocus()
    tImporte.SelStart = 0: tImporte.SelLength = Len(tImporte.Text)
End Sub

Private Sub tImporte_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tImporte.Text) And Trim(tImporte.Text) <> "" Then
            tImporte.Text = Format(tImporte.Text, "#,##0.00")
            Foco tSerie
        End If
    End If
    
End Sub

Private Sub tLClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub tLMemo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = 0 Then tLUsuario.SetFocus
End Sub

Private Sub tLMemo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub tLUsuario_AfterDigit()
    If tLUsuario.UserID <> 0 Then tLClave.SetFocus
End Sub

Private Sub tNumero_Change()
    lFactura.Tag = 0: lFactura.Caption = ""
End Sub

Private Sub tNumero_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Val(lFactura.Tag) <> 0 Or (Trim(tSerie.Text) = "" And Trim(tNumero.Text) = "") Then
            Foco tEnvio
            Exit Sub
        End If
        
        Dim mSerie As String, mNumero As String, mFecha As String
        Dim mID As Long
        mSerie = Trim(tSerie.Text): mNumero = Trim(tNumero.Text)
        
        mID = BuscoDocumento(mSerie, mNumero, mFecha)
        If mID <> 0 Then Foco tEnvio

    End If
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar":
        
        Case "grabar": AccionGrabar
        Case "eliminar": AccionEliminar
        
        Case "cancelar": AccionCancelar
        
        Case "salir": Unload Me
    End Select
    
End Sub

Private Sub tPAutoriza_AfterDigit()
    If tPAutoriza.UserID <> 0 Then AccionGrabar
End Sub

Private Sub tPClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tPAutoriza.SetFocus
End Sub

Private Sub tPMemo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = 0 Then tPClave.SetFocus
End Sub

Private Sub tPMemo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub tSerie_Change()
    lFactura.Tag = 0: lFactura.Caption = ""
End Sub

Private Sub tSerie_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF1
            Call tNumero_KeyPress(vbKeyReturn)
    End Select
    
End Sub

Private Sub tSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(lFactura.Tag) = 0 Then Foco tNumero Else Foco tPMemo
    End If
End Sub


Private Function BuscoDocumento(Optional rSerie As String = "", Optional rNumero As String = "", Optional rFecha As String = "", Optional idDocumento As Long = 0) As Long

    BuscoDocumento = 0
    If Trim(rSerie) = "" And Trim(rNumero) = "" And idDocumento = 0 Then Exit Function
        
    On Error GoTo errBuscaG
    Screen.MousePointer = 11
    
    Dim rIDCliente As Long
    Dim mImporte As Currency
    
    cons = "Select Top 20 DocCodigo, DocTipo, DocCliente, DocSerie as 'Serie', DocNumero as 'Número', DocFecha as Fecha, DocTotal as Importe" & _
                " From Documento " & _
                " Where DocAnulado = 0"
    
    If Trim(rSerie) <> "" Then cons = cons & " And DocSerie = '" & Trim(rSerie) & "'"
    If Trim(rNumero) <> "" Then cons = cons & " And DocNumero = " & Trim(rNumero)
    If idDocumento <> 0 Then cons = cons & " And DocCodigo = " & idDocumento
    
    cons = cons & " Order by DocFecha Desc"
        
    Dim aQ As Integer, aIdSel As Long
    aQ = 0: aIdSel = 0
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        aQ = 1
        aIdSel = rsAux!DocCodigo
        rSerie = Trim(rsAux!serie): rNumero = rsAux(4): rFecha = rsAux!Fecha: rIDCliente = rsAux(2)
        mImporte = rsAux!Importe
        rsAux.MoveNext: If Not rsAux.EOF Then aQ = 2
    End If
    rsAux.Close
        
    Select Case aQ
        Case 0: MsgBox "No hay datos que coincidan con los valores ingresados.", vbExclamation, "No hay datos"
        
        Case 2:
                    Dim miLista As New clsListadeAyuda
                    aIdSel = miLista.ActivarAyuda(cBase, cons, 4000, 3, "Últimas Compras del Cliente")
                    Me.Refresh
                    If aIdSel > 0 Then
                        aIdSel = miLista.RetornoDatoSeleccionado(0)
                        rIDCliente = miLista.RetornoDatoSeleccionado(2)
                        rSerie = miLista.RetornoDatoSeleccionado(3)
                        rNumero = miLista.RetornoDatoSeleccionado(4)
                        rFecha = miLista.RetornoDatoSeleccionado(5)
                        mImporte = miLista.RetornoDatoSeleccionado(6)
                    End If
                    Set miLista = Nothing
    End Select
    
    If aIdSel > 0 Then
        
         cons = "Select Cliente.*, CPeCliente, " & _
                            " (RTrim(isnull(CEmNombre, '')) + ' (' + RTrim(isnull(CEmFantasia, '')) + ')')  as NombreE, " & _
                            " (RTrim(isnull(CPeNombre1, '')) + ' ' + RTrim(isnull(CPeNombre2, '')) + ' ' + RTrim(isnull(CPeApellido1, '')) + ' ' + RTrim(isnull(CPeApellido2, '')))  as NombreP" & _
                    " From Cliente " & _
                        " Left Outer Join CPersona On CliCodigo = CPeCliente " & _
                        " Left Outer Join CEmpresa On CliCodigo = CEmCliente " & _
                    " Where CliCodigo = " & rIDCliente
                    
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            If Not IsNull(rsAux!CPeCliente) Then
                mTexto = Trim(rsAux!NombreP)
            Else
                mTexto = Trim(rsAux!NombreE)
            End If
        End If
        rsAux.Close
               
        tSerie.Text = Trim(UCase(rSerie))
        tNumero.Text = Trim(rNumero)
        lFactura.Caption = Format(rFecha, "(d/mm/yy)") & " " & mTexto
        lFactura.Tag = aIdSel

        tSerie.Tag = Format(rFecha, "dd/mm/yyyy")
        If tSerie.Tag <> dOrigen.Value Then dOrigen.Value = tSerie.Tag
        If tImporte.Text = "" Then tImporte.Text = Format(mImporte, "#,##0.00")
        
        'Busco si hay envio p/documento
        If tEnvio.Enabled Then
            tEnvio.Text = ""
            cons = "Select * from Envio Where EnvDocumento = " & aIdSel
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            If Not rsAux.EOF Then tEnvio.Text = rsAux!EnvCodigo
            rsAux.Close
        End If
        
        BuscoDocumento = aIdSel
    End If
    
    Screen.MousePointer = 0
   
    Exit Function
errBuscaG:
    clsGeneral.OcurrioError "Error al buscar el documento.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub HabilitoCampos(Optional sNuevo As Boolean = False, Optional sEliminar As Boolean = False)

Dim bState As Boolean
Dim bkColor As Long, bkColorNot As Long
            
    bkColor = vbWindowBackground
    bkColorNot = Colores.Inactivo
        
    If Not sNuevo And Not sEliminar Then
        tID.Enabled = True: tID.BackColor = bkColor
    Else
        tID.Enabled = False: tID.BackColor = bkColorNot
    End If
    
    tLFecha.Enabled = False: tLFecha.BackColor = bkColorNot
    
    cDisponibilidad.Enabled = False: cDisponibilidad.BackColor = bkColorNot
    dOrigen.Enabled = False
    cConcepto.Enabled = False: cConcepto.BackColor = bkColorNot
    tImporte.Enabled = False: tImporte.BackColor = bkColorNot
    
    tSerie.Enabled = False: tSerie.BackColor = bkColorNot
    tNumero.Enabled = False: tNumero.BackColor = bkColorNot
    tEnvio.Enabled = False: tEnvio.BackColor = bkColorNot
    tPMemo.Enabled = False: tPMemo.BackColor = bkColorNot
    
    tPUsuario.Enabled = False: tPUsuario.BackColor = bkColorNot
    tPClave.Enabled = False: tPClave.BackColor = bkColorNot
    tPAutoriza.Enabled = False: tPAutoriza.BackColor = bkColorNot
    
    tLMemo.Enabled = False: tLMemo.BackColor = bkColorNot
    tLUsuario.Enabled = False: tLUsuario.BackColor = bkColorNot
    tLClave.Enabled = False: tLClave.BackColor = bkColorNot
    
    If sNuevo Then
        cDisponibilidad.Enabled = True: cDisponibilidad.BackColor = bkColor
        dOrigen.Enabled = True
        cConcepto.Enabled = True: cConcepto.BackColor = bkColor
        tImporte.Enabled = True: tImporte.BackColor = bkColor
        
        tSerie.Enabled = True: tSerie.BackColor = bkColor
        tNumero.Enabled = True: tNumero.BackColor = bkColor
        tEnvio.Enabled = True: tEnvio.BackColor = bkColor
        tPMemo.Enabled = True: tPMemo.BackColor = bkColor
          
        tPUsuario.Enabled = True: tPUsuario.BackColor = bkColor
        tPClave.Enabled = True: tPClave.BackColor = bkColor
        tPAutoriza.Enabled = True: tPAutoriza.BackColor = bkColor
    End If
    
    If sEliminar Then
        tLMemo.Enabled = True: tLMemo.BackColor = bkColor
        tLUsuario.Enabled = True: tLUsuario.BackColor = bkColor
        tLClave.Enabled = True: tLClave.BackColor = bkColor
    End If
    
End Sub

Private Sub AccionNuevo(Optional mDocumento As Long = 0)
    On Error Resume Next
    
    tID.Text = ""
    LimpioFicha
    BuscoCodigoEnCombo cDisponibilidad, paDisponibilidad
    HabilitoCampos sNuevo:=True
    
    Me.Refresh
    Botones False, False, False, True, True, Toolbar1, Me
    
    tImporte.SetFocus
    cDisponibilidad.SetFocus
    BuscoDocumento idDocumento:=mDocumento
    tPUsuario.UserID = paCodigoDeUsuario
    prmNuevo = True
    
End Sub


Private Sub AccionCancelar()
On Error Resume Next
    Screen.MousePointer = 11
    
    Dim mOldID As Long
    mOldID = 0
    If Trim(tID.Text) <> "" Then mOldID = Val(tID.Text)
    
    HabilitoCampos
    LimpioFicha
    
    Botones True, False, False, False, False, Toolbar1, Me
    If mOldID <> 0 Then CargoPendiente mOldID
    
    Foco tID
    
    Screen.MousePointer = 0
    
End Sub

Private Function CargoPendiente(mID As Long)
    Screen.MousePointer = 11
    
Dim mIdDocumento As Long
    
    LimpioFicha
    tID.Text = mID
    
    cons = "Select * From PendientesCaja " & _
                " Where PCaID = " & mID
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    mID = 0: mIdDocumento = 0
    If Not rsAux.EOF Then
        mID = rsAux!PCaID
      
        If Not IsNull(rsAux!PCaDocumento) Then mIdDocumento = rsAux!PCaDocumento
        
        BuscoCodigoEnCombo cDisponibilidad, rsAux!PCaDisponibilidad
        
        dOrigen.Value = rsAux!PCaFPendiente
        BuscoCodigoEnCombo cConcepto, rsAux!PCaConcepto
        tImporte.Text = Format(rsAux!PCaImporte, "#,##0.00")
        
        If Not IsNull(rsAux!PCaEnvio) Then tEnvio.Text = rsAux!PCaEnvio
        
        tPMemo.Text = TextoLargo(rsAux!PCaDetalleOrigen)
        
        tPUsuario.UserID = rsAux!PCaUsrPendiente
        tPAutoriza.UserID = rsAux!PCaUsrAutoriza
        
        lFIngreso.Caption = Format(rsAux!PCaFIngreso, "dd/mm/yyyy hh:mm")
        
        If Not IsNull(rsAux!PCaFLiquidacion) Then tLFecha.Text = Format(rsAux!PCaFLiquidacion, "dd/mm/yyyy")
        If Not IsNull(rsAux!PCaUsrLiquidacion) Then tLUsuario.UserID = rsAux!PCaUsrLiquidacion
        tLMemo.Text = TextoLargo(rsAux!PCaDetalleLiquidacion)
        
        
    End If
    rsAux.Close
    
    If mID <> 0 Then
        'Cargo datos del documento Devolucion ------------------------------------------------------
        sbHelp.Panels("help").Text = "Cargando datos documento ...": sbHelp.Refresh
        
        If mIdDocumento <> 0 Then
            cons = "Select Documento.*, CPeCliente , " & _
                            " (RTrim(isnull(CEmNombre, '')) + ' (' + RTrim(isnull(CEmFantasia, '')) + ')')  as NombreE, " & _
                            " (RTrim(isnull(CPeNombre1, '')) + ' ' + RTrim(isnull(CPeNombre2, '')) + ' ' + RTrim(isnull(CPeApellido1, '')) + ' ' + RTrim(isnull(CPeApellido2, '')))  as NombreP" & _
                        " From Documento, Cliente " & _
                            " left Outer Join CPersona On CliCodigo = CPeCliente " & _
                            " left Outer Join CEmpresa On CliCodigo = CEmCliente " & _
                        " Where DocCliente = CliCodigo " & _
                        " And DocCodigo = " & mIdDocumento
            
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            If Not rsAux.EOF Then
                mTexto = Format(rsAux!DocFecha, "(d/mm/yy)") & " "
                If Not IsNull(rsAux!CPeCliente) Then
                    mTexto = mTexto & Trim(rsAux!NombreP)
                Else
                    mTexto = mTexto & Trim(rsAux!NombreE)
                End If
                
                tSerie.Text = Trim(UCase(rsAux!DocSerie))
                tNumero.Text = Trim(rsAux!DocNumero)
                lFactura.Caption = mTexto
                lFactura.Tag = mIdDocumento
            End If
            rsAux.Close
        End If
        
        sbHelp.Panels("help").Text = ""
        Botones True, True, True, False, False, Toolbar1, Me
        tID.Tag = mID
    End If
    
    Screen.MousePointer = 0
    Exit Function

errCargar:
    clsGeneral.OcurrioError "Error al cargar los datos del registro.", Err.Description
    Screen.MousePointer = 0
    Exit Function
End Function

Private Sub tPUsuario_AfterDigit()
    If tPUsuario.UserID <> 0 Then tPClave.SetFocus
End Sub

Private Sub AccionMenuHelp()
    On Error GoTo errHelp
    Screen.MousePointer = 11
    
    Dim aFile As String
    cons = "Select * from Aplicacion Where AplNombre = 'Pendientes de Caja'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then If Not IsNull(rsAux!AplHelp) Then aFile = Trim(rsAux!AplHelp)
    rsAux.Close
    
    If aFile <> "" Then EjecutarApp aFile
    
    Screen.MousePointer = 0
    Exit Sub
    
errHelp:
    clsGeneral.OcurrioError "Error al activar el archivo de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionEliminar()

    If Val(tID.Tag) = 0 Then Exit Sub
    
    If Trim(tLFecha.Text) <> "" Then
        MsgBox "El pendiente seleccionado ya se liquidó" & vbCrLf & _
                    "Ésta acción permite liquidar pendientes de caja.", vbInformation, "Pendiente Liquidado"
        Exit Sub
    End If

    HabilitoCampos sEliminar:=True
    
    Me.Refresh
    Botones False, False, False, True, True, Toolbar1, Me
        
    tLFecha.Text = Format(gFechaServidor, "dd/mm/yyyy")
    tLMemo.SetFocus
    tLUsuario.UserID = paCodigoDeUsuario
    prmLiquidar = True
    
End Sub

Private Function TextoLargo(mValor As Variant) As String
    On Error Resume Next
    TextoLargo = ""
    TextoLargo = mValor
End Function

Private Sub AyudaPendientes(Optional bLiquidados As Boolean = False)

On Error GoTo errAyuda
    Screen.MousePointer = 11
    
    cons = "Select PCaID as ID, DisNombre as Disponibilidad, PCaFPendiente as 'F/Origen', TPeNombre as Concepto, PCaImporte as Importe, " & _
                        " RTrim(DocSerie) + '-' + cast(DocNumero as char(6)) as Documento"

    If bLiquidados Then
        cons = cons & ", PCaFLiquidacion as Liquidado"
    Else
        cons = cons & ", PCaDetalleOrigen as Comentarios"
    End If
    
    cons = cons & " From PendientesCaja Left Outer Join Documento On PCaDocumento = DocCodigo, " & _
                                    " Disponibilidad, TipoPendiente" & _
                        " Where PCaDisponibilidad = DisID" & _
                        " And PCaConcepto = TPeCodigo"
    
    If Not bLiquidados Then
        cons = cons & " And PCaFliquidacion is Null " & _
                            " Order by PCaFPendiente DESC"
    Else
        cons = cons & " And PCaFLiquidacion is not null" & _
                            " Order by PCaFLiquidacion DESC, PCaFPendiente DESC"
    End If
    
    Dim aIdSel As Long, mW As Currency
    Dim miLista As New clsListadeAyuda
    
    If Not bLiquidados Then mW = 9000 Else mW = 7000
    
    aIdSel = miLista.ActivarAyuda(cBase, cons, mW, 0, "Lista de Pendientes")
    Me.Refresh
    If aIdSel > 0 Then
        aIdSel = miLista.RetornoDatoSeleccionado(0)
    End If
    Set miLista = Nothing
    
    If aIdSel <> 0 Then CargoPendiente aIdSel
    
    Screen.MousePointer = 0
    Exit Sub

errAyuda:
    clsGeneral.OcurrioError "Error al procesar lista de pendientes.", Err.Description
    Screen.MousePointer = 0
End Sub

Public Function fnc_FechaCierreDisp(idDisponibilidad As Long) As Date
On Error GoTo errQuery
Dim rsLoc As rdoResultset

    fnc_FechaCierreDisp = CDate("1/1/1900")
    cons = "Select Top 1 * " & _
                " From SaldoDisponibilidad " & _
                " Where SDiDisponibilidad = " & idDisponibilidad & _
                " Order by SDiFecha DESC"
   
    Set rsLoc = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsLoc.EOF Then
        Dim mHora As String
        fnc_FechaCierreDisp = rsLoc!SDiFecha
        mHora = rsLoc!SDiHora
        
        'Por los saldos iniciales de disponibilidades
        If mHora = "00:00:00" Then fnc_FechaCierreDisp = fnc_FechaCierreDisp - 1
    End If
    rsLoc.Close

errQuery:
End Function


