VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form frmMotivosNoEntrego 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Eliminar envío"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboCamiones 
      Height          =   330
      Left            =   3960
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   840
      Width           =   4455
   End
   Begin VB.CommandButton butCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6960
      TabIndex        =   10
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton butAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   5040
      Width           =   1455
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsMotivos 
      Height          =   2415
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4260
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
   Begin VB.TextBox txtComentario 
      Appearance      =   0  'Flat
      Height          =   1095
      Left            =   1080
      MaxLength       =   150
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1200
      Width           =   7335
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   8685
      TabIndex        =   0
      Top             =   0
      Width           =   8715
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   120
         Picture         =   "frmMotivosNoEntrego.frx":0000
         ScaleHeight     =   450
         ScaleWidth      =   450
         TabIndex        =   1
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Motivos eliminar envío"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cam&ión:"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Suceso:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblEnvio 
      BackStyle       =   0  'Transparent
      Caption         =   "Envío: 8,888.888.88"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
End
Attribute VB_Name = "frmMotivosNoEntrego"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public InfoDetalleMotivos As clsInfMotEntrega
Public prmEnvio As Long
Private colUsuarios As Collection

Private Sub CargoUsuarios()
On Error GoTo errU
Dim oUsu As clsUsuarios
    
    Set colUsuarios = New Collection
    Cons = "SELECT UsuID, UsuIdentificacion FROM Usuarios inner join UsuariosRoles on URoUsuario = UsuID and URoRol = 2 WHERE UsuEstado = 1 Order by UsuIdentificacion"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Set oUsu = New clsUsuarios
        oUsu.ID = RsAux("UsuID")
        oUsu.Identificacion = Trim(RsAux("UsuIdentificacion"))
        colUsuarios.Add oUsu
        RsAux.MoveNext
    Loop
    RsAux.Close
    Exit Sub
errU:
    'clsGeneral.OcurrioError "Error al cargar los usuarios.", Err.Description, "Motivos"
    MsgBox "Error al cargar los usuarios: " & Err.Description, vbCritical, "ATENCIÓN"
End Sub

Private Sub butAceptar_Click()
On Error GoTo errA
    
    Dim I As Integer
    For I = 1 To vsMotivos.Rows - 1
        If (vsMotivos.Cell(flexcpChecked, I, 0) = 1) Then
            If (vsMotivos.Cell(flexcpData, I, 0) = 3 And vsMotivos.Cell(flexcpText, I, 2) = "") Then
                MsgBox "El motivo " & vsMotivos.Cell(flexcpText, I, 1) & " requiere que indique el responsable.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
        End If
    Next
    
    Dim info As New clsInfMotEntrega
    Dim mot As clsMotivoEntrega
    Dim usuDato As Integer

    For I = 1 To vsMotivos.Rows - 1
        If (vsMotivos.Cell(flexcpChecked, I, 0) = 1) Then
            
            If (vsMotivos.Cell(flexcpData, I, 0) = 3 And vsMotivos.Cell(flexcpText, I, 2) = "") Then
                MsgBox "El motivo " & vsMotivos.Cell(flexcpText, I, 1) & " requiere que indique el responsable.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
            
            usuDato = 0
            If (vsMotivos.Cell(flexcpText, I, 2) <> "") Then
                Dim oUsu As clsUsuarios
                For Each oUsu In colUsuarios
                    If (oUsu.Identificacion = vsMotivos.Cell(flexcpText, I, 2)) Then
                        usuDato = oUsu.ID
                        Exit For
                    End If
                Next
            End If
            Set mot = New clsMotivoEntrega
            mot.Motivo = vsMotivos.Cell(flexcpData, I, 1)
            mot.Responsable = usuDato
            info.Motivos.Add mot
            
        End If
    Next
    If (cboCamiones.ListIndex > -1) Then info.Camion = cboCamiones.ItemData(cboCamiones.ListIndex)
    info.Comentario = Trim(txtComentario.Text)
    info.Envio = Me.prmEnvio
    Set Me.InfoDetalleMotivos = info
    Unload Me
    Exit Sub
    
errA:
    objGral.OcurrioError "Error al guardar la información", Err.Description, "Eliminar envío"
End Sub

Private Sub butCancelar_Click()
    Unload Me
End Sub

Private Sub txtComentario_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then vsMotivos.SetFocus
End Sub

Private Sub vsMotivos_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next

    If (Col = 2) Then
        If vsMotivos.Cell(flexcpText, Row, Col) <> "" Then
            vsMotivos.Cell(flexcpChecked, Row, 0) = True
        End If
    ElseIf (Col = 0) Then
        If (vsMotivos.Cell(flexcpChecked, Row, 0) <> 1) Then vsMotivos.Cell(flexcpText, Row, 2) = ""
    End If
    'ArmoDetalle
End Sub

Private Sub vsMotivos_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col = 1)
    If (Col = 2) Then
        If (Val(vsMotivos.Cell(flexcpData, Row, 0)) <> 3) Then
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    Screen.MousePointer = 11
    
    txtComentario.Text = ""
    CargoCombo "SELECT CamID, CamNombre FROM Camiones where Camiones.CamHabilitado = 1", cboCamiones
    
    CargoUsuarios
    
    Dim usuarios As String
    Dim oUsu As clsUsuarios
    For Each oUsu In colUsuarios
        usuarios = usuarios & "|" & oUsu.Identificacion
    Next
    
    With vsMotivos
        .Editable = True
        .Rows = 1
        .Cols = 2
        .Tag = ""
        .FixedCols = 0
        
        .MergeCells = flexMergeFree
        .AllowSelection = False
        .SelectionMode = flexSelectionByRow
        
        .FormatString = " |Motivo|Responsable"
        .ColWidth(0) = 300
        .ColWidth(1) = 1500
        
        .ColComboList(2) = usuarios
        
        .ColDataType(0) = flexDTBoolean
        .ExtendLastCol = True
    
    End With
    CargoMotivosEntregas
    
    Me.lblEnvio.Caption = "Envío: " & prmEnvio
    CargoEnvio
    Screen.MousePointer = 0
    Exit Sub
ErrLoad:
    objGral.OcurrioError "Error al iniciar el formulario.", Err.Description, "Motivos de eliminación"
    Screen.MousePointer = 0
End Sub

Private Sub CargoEnvio()
    Cons = "SELECT EnvCamion FROM Envio WHERE EnvCodigo = " & Me.prmEnvio
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not IsNull(RsAux("EnvCamion")) Then
        BuscoCodigoEnCombo cboCamiones, RsAux("EnvCamion")
    End If
    RsAux.Close
End Sub

Private Sub CargoMotivosEntregas()
On Error GoTo errCME
Dim rsM As rdoResultset
Dim aValor As String
    
    vsMotivos.Rows = 1
    Cons = "SELECT TMETipoEntrega, MenID, MenNombre, IsNull(MenTipoResponsable, 0) MenTipoResponsable " & _
        "FROM MotivosEntrega INNER JOIN TiposMotivosEntregas ON MEnTipo = TMEID " & _
        "WHERE TMEEntidad = 2 AND TMETipoEntrega = 2 ORDER BY MEnNombre"
    Set rsM = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsM.EOF
        vsMotivos.AddItem ""
        vsMotivos.Cell(flexcpData, vsMotivos.Rows - 1, 0) = CStr(rsM("MenTipoResponsable"))
        vsMotivos.Cell(flexcpText, vsMotivos.Rows - 1, 1) = Trim(rsM("MEnNombre"))
        aValor = CStr(rsM("MenID"))
        vsMotivos.Cell(flexcpData, vsMotivos.Rows - 1, 1) = aValor
        rsM.MoveNext
    Loop
    rsM.Close
    Exit Sub
    
errCME:
    objGral.OcurrioError "Error al cargar los motivos de entregas.", Err.Description, "Motivos"
End Sub

