VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRecepcionEnvio 
   BackColor       =   &H00B3DEF5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Envíos"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
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
   ScaleHeight     =   5385
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4200
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecepcionEnvio.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecepcionEnvio.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecepcionEnvio.frx":0224
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tooMenu 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   582
      ButtonWidth     =   1826
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
            Key             =   "save"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Key             =   "undo"
            Object.ToolTipText     =   "Limpiar datos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Envío"
            Key             =   "envio"
            Object.ToolTipText     =   "Ir a formulario de envíos"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C9F1FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   4215
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   17
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox tEnvio 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         MaxLength       =   8
         TabIndex        =   5
         Top             =   120
         Width           =   855
      End
      Begin VB.OptionButton opEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C9F1FF&
         Caption         =   "A Confirmar"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1575
      End
      Begin VB.OptionButton opEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C9F1FF&
         Caption         =   "Nueva Fecha"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton opEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C9F1FF&
         Caption         =   "Entregó"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton opEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C9F1FF&
         Caption         =   "Entregó Parcial"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox tMotivo 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Text            =   "frmRecepcionEnvio.frx":053E
         Top             =   1680
         Width           =   3735
      End
      Begin VB.CheckBox chSendMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00C9F1FF&
         Caption         =   "Enviar mensaje"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2040
         TabIndex        =   12
         Top             =   1440
         Width           =   1695
      End
      Begin VSFlex6DAOCtl.vsFlexGrid vsEntParcial 
         Height          =   975
         Left            =   240
         TabIndex        =   13
         Top             =   2880
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   1720
         _ConvInfo       =   1
         Appearance      =   0
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
         BackColorBkg    =   15794175
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483639
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
         Caption         =   "&Envío:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Motivo:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.OptionButton opGrabar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C9F1FF&
      Caption         =   "&Individual"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   1080
      Width           =   4215
   End
   Begin VB.OptionButton opGrabar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C9F1FF&
      Caption         =   "I&ngresar todo el resto como entregado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   840
      Value           =   -1  'True
      Width           =   4215
   End
   Begin VB.TextBox tCodigo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      MaxLength       =   8
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lCamion 
      BackColor       =   &H007280FA&
      BackStyle       =   0  'Transparent
      Caption         =   "Camión: Martín"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   285
      Left            =   1800
      TabIndex        =   14
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Código:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmRecepcionEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sFEdit As String

Private Sub s_FindEnvio()
On Error GoTo errFE
    Screen.MousePointer = 11
    
    Cons = "Select REvAEntregar, ArtID, ArtCodigo, rTrim(ArtNombre) as ArtNombre  " & _
            " From Envio, RenglonEnvio, Articulo " & _
            " Where EnvCodigo = " & Val(tEnvio.Text) & " And EnvCodImpresion = " & Val(tCodigo.Text) & _
            " And EnvCodigo = RevEnvio And RevArticulo = ArtID And RevAEntregar > 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        MsgBox "No existe un envío pendiente con ese código que pertenezca al código de impresión.", vbExclamation, "Atención"
    Else
        tEnvio.Tag = Val(tEnvio.Text)
        
        'Cargo la lista por si selecciona la opción EntregaParcial.
        Do While Not RsAux.EOF
            
            RsAux.MoveNext
        Loop
    End If
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub
errFE:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al buscar el envío.", Err.Description
End Sub
Private Sub s_GetDatosReparto()
On Error GoTo errGDR
Dim QTotal As Integer, QCamion As Integer
    
    'Busco los datos de la tabla repartoimpresión.
    Screen.MousePointer = 11
    Cons = "Select IsNull(Sum(RReQTotal), 0) as QTotal, IsNull(Sum(RReQCamion), 0) as QCamion, CamNombre, RImModificado, RImCamion " & _
                " From RepartoImpresion, RenglonReparto, Camion" & _
                " Where RImID = " & Val(tCodigo.Text) & " And RImID = RReReparto And RImCamion = CamCodigo" & _
                " Group by CamNombre, RImModificado, RImCamion"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        lCamion.Caption = "Camión: " & Trim(RsAux!CamNombre)
        lCamion.Tag = RsAux!RImCamion
        sFEdit = RsAux!RImModificado
        QTotal = RsAux!QTotal
        QCamion = RsAux!QCamion
        RsAux.Close
        tCodigo.Tag = Trim(tCodigo.Text)
    Else
        RsAux.Close
        Screen.MousePointer = 0
        MsgBox "No se encontraron datos para el código ingresado.", vbExclamation, "Buscar"
        Exit Sub
    End If
    
    If QTotal > QCamion Then
        If QCamion > 0 Then
            MsgBox "Al camión no se le entregó la totalidad de la mercadería, no se podrá dar todo como entregado.", vbExclamation, "Atención"
        Else
            MsgBox "El camión no tiene la mercadería, imposible dar como entregado.", vbExclamation, "Atención"
        End If
    Else
        If QTotal = 0 Then MsgBox "No hay mercadería a entregar.", vbExclamation, "Atención"
    End If

    opGrabar(0).Enabled = (QTotal = QCamion And QCamion > 0)
    opGrabar(0).Value = (QTotal = QCamion And QCamion > 0)
    opGrabar(1).Enabled = QCamion > 0
    opGrabar(1).Value = (Not opGrabar(0).Value And QCamion > 0)

    With tooMenu
        .Buttons("save").Enabled = (QCamion > 0) And opGrabar(0).Enabled
        .Buttons("undo").Enabled = .Buttons("save").Enabled
    End With
    Screen.MousePointer = 0
    Exit Sub
    
errGDR:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al buscar el código de impresión.", Err.Description
End Sub

Private Sub s_SetCtrlEntregaParcial()
    
    With vsEntParcial
        .Enabled = opEstado(3).Value
    End With
    
End Sub

Private Sub s_SetCtrlEstado()
    
    With tFecha
        .Enabled = opEstado(1).Enabled
        If Not .Enabled Then .Text = ""
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
    End With
    
    With tMotivo
        .Enabled = opEstado(0).Value
        .BackColor = IIf(.Enabled, vbWhite, vbButtonFace)
    End With
    With chSendMsg
        .Enabled = opEstado(0).Value
        If Not .Enabled Then .Value = 0 Else .Value = 1
    End With
    
End Sub

Private Sub s_SetCtrlIndividual()
Dim lColor As Long
    
    If opGrabar(1).Enabled And opGrabar(1).Value Then lColor = vbWindowBackground Else lColor = vbButtonFace
        
    With tEnvio
        .Enabled = opGrabar(1).Enabled
        .BackColor = lColor
    End With
    
    opEstado(0).Enabled = opGrabar(1).Value
    opEstado(1).Enabled = opGrabar(1).Value
    opEstado(2).Enabled = opGrabar(1).Value
    opEstado(3).Enabled = opGrabar(1).Value
    
    opEstado(0).Value = False
    opEstado(1).Value = False
    opEstado(2).Value = False
    opEstado(3).Value = False
    
End Sub

Private Sub s_CtrlClean()
    
    sFEdit = ""
    With tooMenu
        .Buttons("save").Enabled = False
        .Buttons("undo").Enabled = False
    End With
    
    opGrabar(0).Value = False
    opGrabar(1).Value = False
    opGrabar(0).Enabled = False
    opGrabar(1).Enabled = False
    
    lCamion.Caption = ""
    
    s_SetCtrlIndividual
    s_SetCtrlEstado
    
    tEnvio.Text = ""
    tMotivo.Text = ""
    
End Sub

Private Sub Form_Load()

    s_CtrlClean
    With vsEntParcial
        .Rows = 0
        .ColWidth(0) = 500
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

End Sub

Private Sub Label1_Click()
    With tCodigo
        If .Enabled Then .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label3_Click()
    With tEnvio
        If .Enabled Then .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label4_Click()
    With tMotivo
        If .Enabled Then .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub opEstado_Click(Index As Integer)
    
    s_SetCtrlEstado
    s_SetCtrlEntregaParcial
    
End Sub

Private Sub opGrabar_Click(Index As Integer)
    s_SetCtrlIndividual
End Sub

Private Sub opGrabar_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 0 Then
'            act_save
        Else
            On Error Resume Next
            tEnvio.SetFocus
        End If
    End If
End Sub

Private Sub tCodigo_Change()
    If Val(tCodigo.Tag) > 0 Then tCodigo.Tag = "": s_CtrlClean
End Sub

Private Sub tCodigo_GotFocus()
    With tCodigo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tCodigo.Tag) > 0 Then
            If opGrabar(0).Enabled Then
                opGrabar(0).SetFocus
            ElseIf opGrabar(1).Enabled Then
                opGrabar(1).SetFocus
            End If
        Else
            s_GetDatosReparto
        End If
    End If
End Sub

Private Sub tEnvio_Change()
    If Val(tEnvio.Tag) > 0 Then s_SetCtrlIndividual
End Sub

Private Sub tEnvio_GotFocus()
    With tEnvio
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tEnvio_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tEnvio.Tag) > 0 Then
            opEstado(0).SetFocus
        Else
            If Not IsNumeric(tEnvio.Text) Then
                MsgBox "No es un código válido.", vbExclamation, "Atención"
            Else
                'Busco el envío.
                s_FindEnvio
            End If
        End If
    End If
End Sub

