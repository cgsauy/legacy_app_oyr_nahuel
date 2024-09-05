VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form frmProveedorService 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Proveedor Service"
   ClientHeight    =   5370
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9660
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProveedorService.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   0
      ScaleHeight     =   4335
      ScaleWidth      =   5895
      TabIndex        =   21
      Top             =   360
      Width           =   5895
      Begin AACombo99.AACombo cClave 
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
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
      Begin VB.TextBox tDatosService 
         Appearance      =   0  'Flat
         BackColor       =   &H000000A0&
         ForeColor       =   &H00FFFFFF&
         Height          =   1605
         Left            =   1080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Text            =   "frmProveedorService.frx":030A
         Top             =   1560
         Width           =   4455
      End
      Begin VB.TextBox tHorario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   3240
         Width           =   2415
      End
      Begin VB.TextBox tService 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   840
         Width           =   4455
      End
      Begin VB.TextBox tTexto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   480
         Width           =   4455
      End
      Begin VB.TextBox tProveedor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   120
         Width           =   4455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "&Clave:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "&Horario:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "&Service:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "&Texto:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Proveedor:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7080
      Top             =   4680
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
            Picture         =   "frmProveedorService.frx":0310
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProveedorService.frx":046A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProveedorService.frx":05C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProveedorService.frx":08DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProveedorService.frx":0A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProveedorService.frx":0B92
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProveedorService.frx":0EAC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   5400
      ScaleHeight     =   3855
      ScaleWidth      =   2055
      TabIndex        =   12
      Top             =   3720
      Width           =   2055
      Begin VB.CommandButton bFiltrar 
         Caption         =   "&Filtrar"
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton bCloseFind 
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   8.25
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1700
         TabIndex        =   19
         Top             =   60
         Width           =   285
      End
      Begin VB.TextBox tFindService 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox tFindTexto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox tFindProveedor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   1920
         X2              =   120
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Todo o parte del"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   60
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Service:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Texto:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nombre Proveedor:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1575
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsList 
      Height          =   3735
      Left            =   7560
      TabIndex        =   11
      Top             =   840
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6588
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483636
      ForeColorFixed  =   -2147483634
      BackColorSel    =   13697023
      ForeColorSel    =   8388608
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   1
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
   Begin MSComctlLib.Toolbar tooMenu 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "edit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "undo"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "find"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "Opciones"
      Begin VB.Menu MnuNew 
         Caption         =   "Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuDelete 
         Caption         =   "Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSave 
         Caption         =   "Grabar"
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuCancel 
         Caption         =   "Cancelar"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuLinea2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "Salir"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmProveedorService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public prm_Proveedor As Long

Private Sub bCloseFind_Click()
    picGrid.Visible = False
    Form_Resize
End Sub

Private Sub bFiltrar_Click()
    Cons = IIf(tFindProveedor.Text <> "", " And (cProv.CEmNombre Like '" & f_GetReplaceQuery(tFindProveedor.Text) & "%' Or cProv.CEmFantasia Like '" & f_GetReplaceQuery(tFindProveedor.Text) & "%')", "")
    Cons = Cons & IIf(tFindTexto.Text <> "", " And (PSeTexto Like '" & f_GetReplaceQuery(tFindTexto.Text) & "%' Or PSeTexto Like '" & f_GetReplaceQuery(tFindTexto.Text) & "%')", "")
    Cons = Cons & IIf(tFindService.Text <> "", " And (Serv.CEmNombre Like '" & f_GetReplaceQuery(tFindService.Text) & "%' Or Serv.CEmFantasia Like '" & f_GetReplaceQuery(tFindService.Text) & "%')", "")
    Cons = f_GetEncabezadoQueryGrid & Cons
    s_FillGrid Cons
End Sub

Private Sub cClave_GotFocus()
On Error Resume Next
    With cClave
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub cClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tHorario.SetFocus
End Sub

Private Sub Form_Load()
'
    ObtengoSeteoForm Me, 100, 100, 5370, 5370
    picDatos.Visible = False
    s_InitGrid
    If prm_Proveedor > 0 Then
        Cons = f_GetEncabezadoQueryGrid & " And CProv.CEmCliente = " & prm_Proveedor
    Else
        Cons = f_GetEncabezadoQueryGrid & " Order By PSeProveedor"
    End If
    s_FillGrid Cons
    s_SetStateEdit False
    If prm_Proveedor > 0 And vsList.Rows = vsList.FixedRows Then
        act_New
        s_LoadFindProveedor
    End If
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
    picGrid.Move 0, tooMenu.Height, picGrid.Width, ScaleHeight - tooMenu.Height
    
    With vsList
        .Move IIf(picGrid.Visible, picGrid.Width + 120, 120), tooMenu.Height + 60, ScaleWidth - IIf(picGrid.Visible, picGrid.Width + 240, 240), ScaleHeight - tooMenu.Height - 120
        If .Visible Then
            .ColWidth(0) = .ClientWidth / 3
            .ColWidth(1) = (.ClientWidth / 3) - IIf(.Rows > .FixedRows, 50, 50)
            .ColWidth(2) = .ClientWidth / 3
        End If
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    GuardoSeteoForm Me
End Sub

Private Sub Label1_Click()
    prj_SetFocus tFindProveedor
End Sub

Private Sub Label10_Click()
    cClave.SetFocus
End Sub

Private Sub Label2_Click()
    prj_SetFocus tFindTexto
End Sub

Private Sub Label3_Click()
    prj_SetFocus tFindService
End Sub

Private Sub Label5_Click()
    prj_SetFocus tProveedor
End Sub

Private Sub Label6_Click()
    prj_SetFocus tTexto
End Sub

Private Sub Label7_Click()
    prj_SetFocus tService
End Sub

Private Sub Label8_Click()
    prj_SetFocus tHorario
End Sub

Private Sub MnuCancel_Click()
    act_Undo
End Sub

Private Sub MnuDelete_Click()
    act_Delete
End Sub

Private Sub MnuExit_Click()
    Unload Me
End Sub

Private Sub MnuNew_Click()
    act_New
End Sub

Private Sub MnuSave_Click()
    act_Save
End Sub

Private Sub tFindProveedor_Change()
    tFindProveedor.Tag = ""
End Sub

Private Sub tFindProveedor_GotFocus()
    With tFindProveedor
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tFindProveedor_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then tFindTexto.SetFocus
    
End Sub

Private Sub tFindService_Change()
    tFindService.Tag = ""
End Sub

Private Sub tFindService_GotFocus()
    With tFindService
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tFindService_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bFiltrar.SetFocus
End Sub

Private Sub tFindTexto_GotFocus()
    With tFindTexto
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tFindTexto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tFindService.SetFocus
End Sub

Private Sub tHorario_GotFocus()
    With tHorario
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tHorario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then act_Save
End Sub

Private Sub tooMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "new": act_New
        Case "edit": act_Edit
        Case "delete": act_Delete
        Case "save": act_Save
        Case "undo": act_Undo
        Case "find": picGrid.Visible = True: Form_Resize
    End Select
End Sub

Private Sub tProveedor_Change()
    tProveedor.Tag = ""
End Sub

Private Sub tProveedor_GotFocus()
    With tProveedor
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tProveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tProveedor.Tag) > 0 Then
            tTexto.SetFocus
        Else
            s_FindEmpresa tProveedor
            Call tProveedor_GotFocus
        End If
    End If
End Sub

Private Sub tService_Change()
    If Val(tService.Tag) > 0 Then
        tDatosService.Text = ""
        tService.Tag = ""
    End If
End Sub

Private Sub tService_GotFocus()
    With tService
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tService_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tService.Tag) > 0 Then
            cClave.SetFocus
        Else
            s_FindEmpresa tService
            If Val(tService.Tag) > 0 Then
                s_GetInfoService
                'Cargo todos los nombres de las direcciones auxiliares.
                db_FillComboService
            End If
            Call tService_GotFocus
        End If
    End If
End Sub

Private Sub tTexto_GotFocus()
    With tTexto
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tTexto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tService.SetFocus
End Sub

Private Sub s_SetStateEdit(ByVal bEdit As Boolean)
    With tooMenu
        .Buttons("new").Enabled = Not bEdit
        .Buttons("edit").Enabled = Not bEdit And vsList.Row >= vsList.FixedRows
        .Buttons("delete").Enabled = Not bEdit And vsList.Row >= vsList.FixedRows
        .Buttons("save").Enabled = bEdit
        .Buttons("undo").Enabled = bEdit
        .Buttons("find").Enabled = Not bEdit
    End With
    vsList.Visible = Not bEdit
    picGrid.Visible = Not bEdit
    picDatos.Visible = bEdit
    Call Form_Resize
End Sub

Private Sub s_InitGrid()
    
    With vsList
        .Rows = 1: .Cols = 1
        .FormatString = "Proveedor|Texto|Service|Horario"
        .ColHidden(3) = True
    End With
    
End Sub

Private Sub s_CleanCtrlEdit()
    picDatos.Tag = ""
    With tProveedor
        .Text = "": .Tag = ""
    End With
    tTexto.Text = ""
    With tService
        .Text = "": .Tag = ""
    End With
    tHorario.Text = ""
    tDatosService.Text = ""
    cClave.Clear
End Sub
Private Sub act_New()
On Error Resume Next
    s_CleanCtrlEdit
    s_SetStateEdit True
    tProveedor.SetFocus
End Sub

Private Sub act_Edit()
        
    s_CleanCtrlEdit
    With vsList
        picDatos.Tag = Val(.Cell(flexcpData, .Row, 0))
        tProveedor.Text = Trim(.Cell(flexcpText, .Row, 0))
        tProveedor.Tag = .Cell(flexcpData, .Row, 1)
        tTexto.Text = .Cell(flexcpText, .Row, 1)
        tService.Text = .Cell(flexcpText, .Row, 2)
        tService.Tag = .Cell(flexcpData, .Row, 2)
        If Val(tService.Tag) > 0 Then s_GetInfoService
        db_FillComboService
        cClave.Text = .Cell(flexcpData, .Row, 3)
        tHorario.Text = .Cell(flexcpText, .Row, 3)
    End With
    s_SetStateEdit True
    tProveedor.SetFocus
End Sub

Private Sub act_Delete()
On Error GoTo errAD
    
    If vsList.Row > vsList.FixedRows Then
        If MsgBox("¿Confirma eliminar el registro seleccionado?", vbQuestion + vbYesNo, "Atención") = vbNo Then Exit Sub
        Screen.MousePointer = 11
        Cons = "Select * From ProveedorService Where PSeID = " & vsList.Cell(flexcpData, vsList.Row, 0)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.Delete
        RsAux.Close
        vsList.RemoveItem vsList.Row
    End If
    s_SetStateEdit False
    Screen.MousePointer = 0
    Exit Sub
    
errAD:
    objGral.OcurrioError "Error al intentar eliminar"
End Sub

Private Sub act_Save()
On Error GoTo errAS
Dim lIDNew As Long

    If MsgBox("¿Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "Grabar") <> vbYes Then Exit Sub
    Screen.MousePointer = 11
    Cons = "Select * From ProveedorService Where PSeID = " & Val(picDatos.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.AddNew
    Else
        RsAux.Edit
    End If
    RsAux!PSeProveedor = Val(tProveedor.Tag)
    RsAux!PSeTexto = IIf(Trim(tTexto.Text) <> "", Trim(tTexto.Text), Null)
    RsAux!PSeService = IIf(Val(tService.Tag) > 0, Val(tService.Tag), Null)
    RsAux!PSeHorario = IIf(Trim(tHorario.Text) <> "", tHorario.Text, Null)
    RsAux!PSeClave = IIf(Trim(cClave.Text) <> "", cClave.Text, Null)
    RsAux.Update
    RsAux.Close
    
    If Val(picDatos.Tag) = 0 Then
        Cons = "Select Max(PSeID) From ProveedorService Where PSeProveedor = " & Val(tProveedor.Tag) & _
                    " And PSeTexto " & IIf(Trim(tTexto.Text) <> "", " = '" & Trim(tTexto.Text) & "'", "Is Null") & _
                    " And PSeService " & IIf(Val(tService.Tag) > 0, " = " & Val(tService.Tag), " Is Null")
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then lIDNew = RsAux(0)
        RsAux.Close
        
        If lIDNew > 0 Then
            s_FillGrid f_GetEncabezadoQueryGrid & " Order By PSeProveedor"
        End If
    Else
        With vsList
            .Cell(flexcpText, .Row, 0) = tProveedor.Text
            .Cell(flexcpData, .Row, 1) = Val(tProveedor.Tag)
            
            'Texto
            .Cell(flexcpText, .Row, 1) = Trim(tTexto.Text)
            
            'Service
            .Cell(flexcpText, .Row, 2) = Trim(tService.Text)
            .Cell(flexcpData, .Row, 2) = Val(tService.Tag)
            
            'Horario
            .Cell(flexcpText, .Row, 3) = Trim(tHorario.Text)
        End With
    End If
    act_Undo
    Screen.MousePointer = 0
    Exit Sub
errAS:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al grabar la información.", Err.Description
End Sub

Private Sub act_Undo()
    picDatos.Tag = ""
    s_SetStateEdit False
    vsList.SetFocus
End Sub

Private Sub s_FindEmpresa(ByRef tText As TextBox)
On Error GoTo errFE
Dim sFind As String
    
    Screen.MousePointer = 11
    
    sFind = "'" & f_GetReplaceQuery(tText.Text) & "%'"
    Cons = "Select CEmCliente, CEmNombre as 'Nombre', CEmFantasia as 'Fantasia' From CEmpresa Where CEmNombre Like " & sFind & " or CEmFantasia Like " & sFind
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        RsAux.MoveNext
        If RsAux.EOF Then
            RsAux.MoveFirst
            With tText
                .Text = Trim(RsAux(1))
                .Tag = RsAux(0)
            End With
        Else
            Dim objLista As New clsListadeAyuda
            If objLista.ActivarAyuda(cBase, Cons, 6000, 1, "Empresas") > 0 Then
                With tText
                    .Text = objLista.RetornoDatoSeleccionado(1)
                    .Tag = objLista.RetornoDatoSeleccionado(0)
                End With
            End If
            Set objLista = Nothing
        End If
    Else
        MsgBox "No existen empresas para el dato ingresado.", vbInformation, "Atención"
    End If
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub
errFE:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al buscar la empresa.", Err.Description
End Sub

Private Sub s_GetInfoService()
On Error GoTo errGIS
Dim lIDDir As Long
Dim sTelef As String

    Screen.MousePointer = 11
    Cons = "Select * From Cliente Where CliCodigo = " & Val(tService.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!CliDireccion) Then lIDDir = RsAux!CliDireccion
    End If
    RsAux.Close
    If lIDDir > 0 Then tDatosService.Text = " " & objGral.ArmoDireccionEnTexto(cBase, lIDDir, entrecalles:=True) & vbCrLf
    
    sTelef = ""
    Cons = "Select TTeNombre, TelNumero, IsNull(TelInterno, '') as TelInterno From Telefono, TipoTelefono " & _
                " Where TelCliente = " & Val(tService.Tag) & " And TelTipo = TTecodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        If sTelef <> "" Then sTelef = sTelef & vbCrLf
        sTelef = sTelef & " · " & Trim(RsAux!TTeNombre) & " " & _
                objGral.RetornoFormatoTelefono(cBase, Trim(RsAux!TelNumero), lIDDir) & IIf(RsAux!TelInterno <> "", "  " & RsAux!TelInterno, "")
        RsAux.MoveNext
    Loop
    RsAux.Close
    If sTelef <> "" Then
        If tDatosService.Text <> "" Then tDatosService.Text = tDatosService.Text & vbCrLf & "Teléfonos: " & vbCrLf
    End If
    tDatosService.Text = tDatosService.Text & sTelef
    Screen.MousePointer = 0
    Exit Sub
    
errGIS:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al buscar la información del service.", Err.Description
End Sub

Private Sub s_FillGrid(ByVal sSQl As String)
On Error GoTo errFG
Dim lID As Long
Dim sClave As String

    Screen.MousePointer = 11
    vsList.Rows = 1
    Set RsAux = cBase.OpenResultset(sSQl, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        With vsList
            .AddItem Trim(RsAux(2))
            lID = RsAux(0): .Cell(flexcpData, .Rows - .FixedRows, 0) = lID
            
            'Texto
            .Cell(flexcpText, .Rows - .FixedRows, 1) = Trim(RsAux(3))
            
            lID = RsAux(1): .Cell(flexcpData, .Rows - .FixedRows, 1) = lID
            
            'Service
            .Cell(flexcpText, .Rows - .FixedRows, 2) = Trim(RsAux(5))
            lID = RsAux(4): .Cell(flexcpData, .Rows - .FixedRows, 2) = lID
            
            'Horario
            .Cell(flexcpText, .Rows - .FixedRows, 3) = Trim(RsAux(6))
            
            sClave = Trim(RsAux("PSeClave"))
            .Cell(flexcpData, .Rows - .FixedRows, 3) = sClave
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub
errFG:
    objGral.OcurrioError "Error al cargar la información en la grilla.", Err.Description
End Sub

Private Function f_GetEncabezadoQueryGrid() As String

    f_GetEncabezadoQueryGrid = _
                        "Select PSeID, PSeProveedor, cProv.CEmNombre, IsNull(PSeTexto, '') as PSeTexto, IsNull(PSeService, 0) as PSeService, IsNull(Serv.CEmNombre, '') as SerNombre, IsNull(PSeHorario, '') as PSeHorario, IsNull(PSeClave, '') as PSeClave " & _
                        " From ProveedorService " & _
                                " Left Outer Join CEmpresa Serv On Serv.CEmCliente = PSeService " & _
                        ", CEmpresa as CProv" & _
                        " Where PSeProveedor = CProv.CEmCliente"
End Function

Private Function f_GetReplaceQuery(ByVal sText As String) As String
    f_GetReplaceQuery = Replace(Replace(sText, " ", "%"), "'", "%")
End Function

Private Sub s_LoadFindProveedor()
On Error GoTo errLFP
    Cons = "Select CEmCliente, CEmNombre as 'Nombre', CEmFantasia as 'Fantasia' From CEmpresa Where CEmCliente = " & prm_Proveedor
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        With tProveedor
            .Text = Trim(RsAux(1))
            .Tag = RsAux(0)
        End With
    End If
    RsAux.Close
    Exit Sub

errLFP:
    objGral.OcurrioError "Error al buscar el proveedor.", Err.Description, "Error (loadfindproveedor)"

End Sub

Private Sub db_FillComboService()
On Error GoTo errFCS
Dim rsS As rdoResultset
Dim bIn As Boolean
Dim sChange As String
    
    Screen.MousePointer = 11
    sChange = cClave.Text
    cClave.Clear
    bIn = False
    Cons = "Select Distinct(DAuNombre) From DireccionAuxiliar Where DAuCliente = " & Val(tService.Tag) & _
                " And DAuNombre Is Not Null"
    Set rsS = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsS.EOF
        cClave.AddItem Mid(Trim(rsS(0)), 1, 10)
        If LCase(Mid(Trim(rsS(0)), 1, 10)) = "service" Then bIn = True
        rsS.MoveNext
    Loop
    rsS.Close
    If Not bIn Then cClave.AddItem "Service"
    
    If Val(picDatos.Tag) > 0 Then
        If Trim(sChange) <> "" Then
            cClave.Text = sChange
        Else
            cClave.Text = "Service"
        End If
    Else
        cClave.Text = "Service"
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errFCS:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al cargar el combo de claves.", Err.Description
End Sub


