VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMercaAReclamar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Retornar envío impreso"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7155
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
   ScaleHeight     =   6015
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picDatos 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1575
      ScaleWidth      =   7095
      TabIndex        =   14
      Top             =   480
      Width           =   7095
      Begin VB.TextBox tMotivo 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Text            =   "frmMercaAReclamar.frx":0000
         Top             =   1080
         Width           =   6975
      End
      Begin VB.CheckBox chSendMsg 
         Appearance      =   0  'Flat
         Caption         =   "Enviar mensaje"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1800
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox cbCombo 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox cHora 
         Height          =   315
         Left            =   5280
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lbTitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Condiciones para el nuevo estado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Width           =   6735
      End
      Begin VB.Label lbfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lbMemo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Comentario:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lbCombo 
         BackStyle       =   0  'Transparent
         Caption         =   "&Estado:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lbHora 
         BackStyle       =   0  'Transparent
         Caption         =   "&Hora"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4800
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   540
      Left            =   2400
      TabIndex        =   13
      Top             =   4440
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   953
      ButtonWidth     =   2037
      ButtonHeight    =   953
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
            Key             =   "save"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Devuelve todo"
            Key             =   "devuelve"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Retiene todo"
            Key             =   "retiene"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "exit"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsArticulos 
      Height          =   2175
      Left            =   0
      TabIndex        =   9
      Top             =   2160
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3836
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483639
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   4
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   2
      RowHeightMin    =   255
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5520
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMercaAReclamar.frx":0006
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMercaAReclamar.frx":0118
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMercaAReclamar.frx":046A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMercaAReclamar.frx":08BC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbDireccion 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Av Italia 2545/604"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Envío: 10456"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lbMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Dividir un envío que está en entrega se utiliza para dejar los artículos que no fueron entregados al cliente en un nuevo envío"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   5280
      Width           =   6735
   End
   Begin VB.Shape shfac 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      FillColor       =   &H00DCFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   7020
   End
End
Attribute VB_Name = "frmMercaAReclamar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public prmInvocacion As Byte    '0) cambio fecha 1) anular el envío, 2) cambia camión
Public prmEnvio As Long

Private Sub loc_SetGridDevRet(ByVal bDevuelve As Boolean)
Dim iQ As Integer
Dim iCol As Byte, iCol0 As Byte
    If bDevuelve Then iCol0 = 4: iCol = 5 Else iCol = 4: iCol0 = 5
    
    With vsArticulos
        For iQ = .FixedRows To .Rows - 1
            If Val(.Cell(flexcpText, iQ, iCol0)) > 0 Then
                .Cell(flexcpText, iQ, iCol) = Val(.Cell(flexcpText, iQ, iCol)) + Val(.Cell(flexcpText, iQ, iCol0))
                .Cell(flexcpText, iQ, iCol0) = 0
            End If
            If Val(.Cell(flexcpText, iQ, iCol)) > 0 Then .Cell(flexcpBackColor, iQ, iCol) = &HADDEFF '&H66CCFF
        Next
    End With
End Sub
Private Sub loc_SetColorNormal(ByVal bDevuelve As Boolean)
Dim iQ As Integer
Dim iCol As Integer
    If bDevuelve Then iCol = 5 Else iCol = 4
    With vsArticulos
        For iQ = .FixedRows To .Rows - 1
            If .Cell(flexcpBackColor, iQ, iCol) = &HADDEFF Then
                .Cell(flexcpBackColor, iQ, iCol) = vbWindowBackground
            End If
        Next
    End With
End Sub
Private Sub loc_DBDevuelve()
On Error GoTo errInit
    'pongo todo como devuelto en la lista.
    loc_SetGridDevRet True
    
    'pregunto
    If MsgBox("¿Confirma grabar la información?" & vbCrLf & vbCrLf & "Si desea puede validar en la grilla las cantidades ajustadas que devuelve el camión.", vbQuestion + vbYesNo, "Devolver toda la mercadería") = vbYes Then
        
    Else
        loc_SetColorNormal
    End If
errInit:
End Sub

Private Sub loc_FindEnvio()
On Error GoTo errFE
Dim lAux As Long
Dim iCodImpresion As Integer

    Screen.MousePointer = 11
    Toolbar1.Buttons("save").Enabled = False
    vsArticulos.Rows = 1
    Cons = "Select EnvCodigo, IsNull(EnvVaCon, 0) as VaCon, EnvEstado, EnvFModificacion, EnvDireccion, EnvCodImpresion" & _
                " From Envio " & _
                " Where ((EnvCodigo = " & prmEnvio & " And EnvVaCon Is Null) " & _
                        "Or (Abs(EnvVaCon) IN (Select abs(EnvVaCon) From Envio Where EnvCodigo = " & prmEnvio & " And EnvVaCon Is Not Null)))"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        RsAux.Close
        Screen.MousePointer = 0
        MsgBox "No existe un envío con ese código.", vbExclamation, "Atención"
        Exit Sub
    Else
        iCodImpresion = RsAux("EnvCodImpresion")
        If RsAux("EnvEstado") <> 3 Then
            Screen.MousePointer = 0
            RsAux.Close
            MsgBox "El envío no tiene el estado impreso, para modificarlo acceda al formulario de envíos.", vbExclamation, "Atención"
            Exit Sub
        Else
            If RsAux("VaCon") <> 0 Then
                MsgBox "El envío en un VA CON.", vbInformation, "Atención"
            End If
            lbDireccion.Caption = objGral.ArmoDireccionEnTexto(cBase, RsAux("EnvDireccion"))
            vsArticulos.Tag = RsAux("EnvFModificacion")
        End If
        RsAux.Close
    End If
        
    Cons = "Select Sum(REvAEntregar) as QArt, Sum(ReECantidadTotal) as QT, Sum(ReECantidadEntregada) as QE, ArtID, ArtCodigo, rTrim(ArtNombre) as ArtNombre " & _
            " From RenglonEnvio, Articulo, RenglonEntrega " & _
            " Where REvEnvio = " & prmEnvio & _
            " And RevArticulo = ArtID And RevAEntregar > 0 And ReEArticulo = ArtID And ReECodImpresion = " & iCodImpresion & _
            " Group by ArtID, ArtCodigo, ArtNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    'Cargo la lista por si selecciona la opción EntregaParcial.
    Do While Not RsAux.EOF
        With vsArticulos
            .AddItem "(" & Format(RsAux!ArtCodigo, "000,000") & ") " & Trim(RsAux!ArtNombre)
            .Cell(flexcpText, .Rows - 1, 1) = RsAux("QArt")
            .Cell(flexcpText, .Rows - 1, 2) = RsAux("QT")
            .Cell(flexcpText, .Rows - 1, 3) = RsAux("QE")
            If RsAux("QE") = RsAux("QT") Then
                'El camión tiene o no tienen toda la mercadería por lo tanto devuelve todo
                If RsAux("QE") > 0 Then
                    .Cell(flexcpText, .Rows - 1, 4) = 0
                    .Cell(flexcpText, .Rows - 1, 5) = RsAux("QArt")
                Else
                    .Cell(flexcpText, .Rows - 1, 4) = 0
                    .Cell(flexcpText, .Rows - 1, 5) = 0
                    .Cell(flexcpBackColor, .Rows - 1, 4, , 5) = &HE0E0E0
                End If
            Else
                'El camión tiene asignada parte de la mercadería.
                'Por lo tanto siempre le voy a restar al camión.
                If RsAux("QE") > RsAux("QArt") Then
                    .Cell(flexcpText, .Rows - 1, 4) = 0
                    .Cell(flexcpText, .Rows - 1, 5) = RsAux("QArt")
                Else
                    .Cell(flexcpText, .Rows - 1, 4) = RsAux("QArt") - RsAux("QE")
                    .Cell(flexcpText, .Rows - 1, 5) = RsAux("QE")
                End If
            End If
            .Cell(flexcpBackColor, .Rows - 1, 0, , 3) = vbWindowBackground
            .Cell(flexcpBackColor, .Rows - 1, 1) = &HFFF5F0 '14857624
            .Cell(flexcpBackColor, .Rows - 1, 3) = &HFFF5F0

            lAux = RsAux!ArtID
            .Cell(flexcpData, .Rows - 1, 0) = lAux
            lAux = RsAux("QArt"): .Cell(flexcpData, .Rows - 1, 1) = lAux
            lAux = RsAux("QT"): .Cell(flexcpData, .Rows - 1, 2) = lAux
            lAux = RsAux("QE"): .Cell(flexcpData, .Rows - 1, 3) = lAux
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    Toolbar1.Buttons("save").Enabled = (vsArticulos.Rows > 1)
    On Error Resume Next
    If vsArticulos.Rows > 1 Then vsArticulos.SetFocus
    Screen.MousePointer = 0
    Exit Sub
errFE:
    Screen.MousePointer = 0
    vsArticulos.Rows = 1
    objGral.OcurrioError "Error al buscar el envío.", Err.Description
End Sub

Private Sub actSave()
On Error GoTo errSave
Dim iQ As Integer
Dim bQuedan As Boolean, bHay As Boolean
Dim rsEnv As rdoResultset, rsNew As rdoResultset
    
    lbMsg.Caption = "Almacenando"
    
    With vsArticulos
        For iQ = 1 To .Rows - 1
            If Val(.Cell(flexcpText, iQ, 0)) <> Val(.Cell(flexcpText, iQ, 2)) Then
                bQuedan = True
            End If
            If Val(.Cell(flexcpText, iQ, 2)) > 0 Then bHay = True
            If bHay And bQuedan Then Exit For
        Next
    End With
        Exit Sub
        
errBT:
    Screen.MousePointer = vbDefault
    objGral.OcurrioError "Error al intentar iniciar la transacción para dividir el envío.", Err.Description, "Dividir envíos"
    Exit Sub

errorET:
    Resume ErrTransaccion
    Exit Sub
    
ErrTransaccion:
    Screen.MousePointer = vbDefault
    cBase.RollbackTrans
    objGral.OcurrioError "Error al grabar cuando se dividía el envío.", Err.Description, "Dividir envíos"
    Exit Sub
    
errSave:
    objGral.OcurrioError "Error al intentar al dividir el envío.", Err.Description, "Dividir envíos"
End Sub

Private Sub Form_Load()

    picDatos.Visible = (Me.prmInvocacion <> 1)
    Select Case prmInvocacion
        Case 0
            With cbCombo
                .Clear
                .AddItem "A confirmar"
                .AddItem "Nueva fecha"
                .ListIndex = 1
            End With
            lbCombo.Caption = "&Estado:"
            Me.Height = 6345
            Me.Caption = "Cambiar fecha a envío"
            cHora.Clear
            tMotivo.Text = ""
            
        Case 1
            vsArticulos.Top = picDatos.Top
            Me.Height = 4750
            picDatos.Height = 0
            Me.Caption = "Anular envío"
        
        Case 2
            Me.Caption = "Cambiar camión"
            Me.Height = 6345 - 840
            picDatos.Height = 735
            lbfecha.Visible = False
            tFecha.Visible = False
            lbHora.Visible = False
            cHora.Visible = False
            lbTitulo.Caption = "Cambio de camionero"
            lbCombo.Caption = "&Camión:"
            CargoCombo "Select CamCodigo, CamNombre From Camion Order By CamNombre", cbCombo
    End Select
    
    vsArticulos.Top = picDatos.Top + picDatos.Height + 120
    Toolbar1.Top = vsArticulos.Top + vsArticulos.Height + 120
    shfac.Top = Toolbar1.Top + Toolbar1.Height + 120
    lbMsg.Top = shfac.Top + 120
    
    With vsArticulos
        .Rows = 1
        .FixedRows = 1
        .FormatString = "Artículo|Q Env|Q CImp|Entregada|Retiene|>Devuelve"
        .FixedCols = 4
        .RowHeight(0) = 315
        .ColWidth(0) = 3400
        .BackColorSel = vbInfoBackground
        .ForeColorSel = vbWindowText
        .FocusRect = flexFocusHeavy
    End With
    Toolbar1.Buttons("save").Enabled = False
    If prmEnvio > 0 Then loc_FindEnvio
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    vsArticulos.Left = 60
    vsArticulos.Width = ScaleWidth - 120
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "save": actSave
        Case "devuelve": loc_DBDevuelve
        Case "retiene"
        Case "exit": Unload Me
    End Select
End Sub

Private Sub vsArticulos_GotFocus()
    lbMsg.Caption = "Seleccione la columna e ingrese la cantidad de artículos que retiene o devuelve el camionero. (+ o - suma o resta). Las filas en gris no puede modificarlas."
End Sub

Private Sub vsArticulos_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errkd
Dim iC1 As Byte, iC2 As Byte
    If Shift <> 0 Then Exit Sub
    With vsArticulos
        If .Cell(flexcpBackColor, .Row, 5) = &HE0E0E0 Then Exit Sub
        Select Case KeyCode
            Case vbKeyAdd
                'Dada la columna que es resto a la otra.
                iC1 = .Col
                If .Col = 5 Then iC2 = 4 Else iC2 = 5
                If Val(.Cell(flexcpText, .Row, iC2)) > 0 Then
                    .Cell(flexcpText, .Row, iC1) = Val(.Cell(flexcpText, .Row, iC1)) + 1
                    .Cell(flexcpText, .Row, iC2) = Val(.Cell(flexcpText, .Row, iC2)) - 1
                End If
            Case vbKeySubtract
                'Dada la columna que es sumo a la otra.
                iC1 = .Col
                If .Col = 5 Then iC2 = 4 Else iC2 = 5
                If Val(.Cell(flexcpText, .Row, iC1)) > 0 Then
                    .Cell(flexcpText, .Row, iC1) = Val(.Cell(flexcpText, .Row, iC1)) - 1
                    .Cell(flexcpText, .Row, iC2) = Val(.Cell(flexcpText, .Row, iC2)) + 1
                End If
        End Select
    End With
errkd:
End Sub

Private Sub vsArticulos_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn And Toolbar1.Buttons("save").Enabled Then actSave
End Sub
