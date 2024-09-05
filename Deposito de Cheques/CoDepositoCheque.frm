VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form CoDepositoCheque 
   BackColor       =   &H8000000B&
   Caption         =   "Depósito de Cheques Diferidos"
   ClientHeight    =   6135
   ClientLeft      =   2370
   ClientTop       =   2370
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CoDepositoCheque.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   9000
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
      ForeColor       =   &H00000080&
      Height          =   1515
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   8775
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   4680
         TabIndex        =   11
         Top             =   705
         Width           =   735
         _ExtentX        =   1296
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
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   705
         Width           =   2055
      End
      Begin VB.CommandButton bConsultar 
         Caption         =   "&Consultar"
         Height          =   350
         Left            =   7680
         TabIndex        =   5
         Top             =   1065
         Width           =   975
      End
      Begin AACombo99.AACombo cCondicion 
         Height          =   315
         Left            =   6480
         TabIndex        =   12
         Top             =   705
         Width           =   1515
         _ExtentX        =   2672
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de &Vencimiento:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   730
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Moneda:"
         Height          =   255
         Left            =   4020
         TabIndex        =   9
         Top             =   730
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"CoDepositoCheque.frx":0442
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   180
         Width           =   8655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Condición:"
         Height          =   255
         Left            =   5640
         TabIndex        =   7
         Top             =   730
         Width           =   915
      End
   End
   Begin VB.CommandButton bGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3840
      TabIndex        =   3
      Top             =   5625
      Width           =   1095
   End
   Begin ComctlLib.ListView lLista 
      Height          =   3975
      Left            =   60
      TabIndex        =   0
      Top             =   1620
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Vencimiento"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Banco Emisor"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Cheque"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Importe"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Librado"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Banco a Depositar"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label lTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6840
      TabIndex        =   2
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label lTTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total a Depositar:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   5040
      TabIndex        =   1
      Top             =   5640
      Width           =   3855
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CoDepositoCheque.frx":052A
            Key             =   "check"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CoDepositoCheque.frx":0844
            Key             =   "nocheck"
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuBDerecho 
      Caption         =   "BotonDerecho"
      Visible         =   0   'False
      Begin VB.Menu MnuVerSucursal 
         Caption         =   "&Editar Sucursal a Depositar"
      End
      Begin VB.Menu MnuVerCheque 
         Caption         =   "Información del &Cheque"
      End
      Begin VB.Menu MnuVerL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVerDeuda 
         Caption         =   "&Deuda en Cheques"
      End
      Begin VB.Menu MnuVerL2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVerCancelar 
         Caption         =   "Cancelar"
      End
   End
End
Attribute VB_Name = "CoDepositoCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim aTexto As String

Private Sub bConsultar_Click()

    'Valido el ingreso de los campos para realizar la consulta.------------------------------------------
    aTexto = ValidoPeriodoFechas(tFecha.Text) ', True)
    If aTexto = "" Then
        MsgBox "El período de fechas ingresado no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tFecha
        Exit Sub
    End If
    If cMoneda.ListIndex = -1 Then
        MsgBox "Debe seleccionar una moneda para realizar la consulta.", vbExclamation, "ATENCIÓN"
        Foco cMoneda
        Exit Sub
    End If
    '------------------------------------------------------------------------------------------------------------
    
    CargoCheques Trim(tFecha.Text)
    If lLista.ListItems.Count > 0 Then lLista.SetFocus
    
End Sub

Private Sub bGrabar_Click()

    'Valido los datos
    If CCur(lTotal.Caption) = 0 Then MsgBox "Debe seleccionar los cheques a depositar.", vbExclamation, "ATENCIÓN": Exit Sub
    
    If MsgBox("Confirma grabar el depósito de los cheques seleccionados.", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    
    AccionGrabar
    
End Sub


Private Sub cMoneda_Change()
    Selecciono cMoneda, cMoneda.Text, gTecla
End Sub

Private Sub cMoneda_GotFocus()
    cMoneda.SelStart = 0
    cMoneda.SelLength = Len(cMoneda.Text)
End Sub

Private Sub cMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
    gTecla = KeyCode
    gIndice = cMoneda.ListIndex
End Sub

Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    cMoneda.ListIndex = gIndice
    If KeyAscii = vbKeyReturn Then oADepositar.SetFocus
End Sub

Private Sub cMoneda_KeyUp(KeyCode As Integer, Shift As Integer)
    ComboKeyUp cMoneda
End Sub

Private Sub cMoneda_LostFocus()
    gIndice = -1
    cMoneda.SelLength = 0
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    Screen.MousePointer = 11
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    'SetearLView lvValores.Grilla Or lvValores.FullRow, lLista
    tFecha.Text = Format(Now, "d/mm/yyyy")
   
    'Cargo las monedas en el combo-------------------------
    Cons = "Select MonCodigo, MonSigno from Moneda where MonFactura = 1 Order by MonSigno"
    CargoCombo Cons, cMoneda, ""
    '--------------------------------------------------------------
   
End Sub


Private Sub CargoCheques(fecha As String)

Dim RsSuc As rdoResultset
Dim aSucCodigo As Long, aSucNombre As String        'Para Depositar
Dim aTotal As Currency

    On Error GoTo errPago
    Screen.MousePointer = 11
    lLista.ListItems.Clear
    bGrabar.Enabled = False
    aSucCodigo = 0
    aTotal = 0
    
    'Armo la Consulta de Cheques-------------------------------------------------------------------------
    Cons = "Select * from ChequeDiferido, SucursalDeBanco, BancoSSFF" _
            & " Where CDiMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
            & ConsultaDeFecha("And", "CDiVencimiento", fecha) _
            & " And CDiSucursal = SBaCodigo" _
            & " And CDiBanco = BanCodigo"
    
    If oADepositar.Value Then
        Cons = Cons & " And CDiCobrado = NULL"
        bConsultar.Tag = "AD"
        If lLista.ColumnHeaders.Count = 7 Then lLista.ColumnHeaders.Remove 7
    Else
        Cons = Cons & " And CDiCobrado <> NULL"
        bConsultar.Tag = "DE"
        If lLista.ColumnHeaders.Count = 6 Then lLista.ColumnHeaders.Add 7, , "Depositado", 700, 2
    End If
    
    Cons = Cons & " Order by CDiVencimiento"
    '-----------------------------------------------------------------------------------------------------------
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    Do While Not RsAux.EOF
        Set itmX = lLista.ListItems.Add(, "A" & RsAux!CDiCodigo, Format(RsAux!CDiVencimiento, "dd/mm/yy"))
        'TAG: CxxxxxSxxxx  (Cliente, Sucursal)
        itmX.Tag = "C" & RsAux!CDiCliente
        
        itmX.SmallIcon = "check"
        itmX.SubItems(1) = Trim(RsAux!BanNombre) & " (" & Trim(RsAux!SBaNombre) & ")"
        
        itmX.SubItems(2) = Trim(RsAux!CDiSerie) & " " & Trim(RsAux!CDiNumero)
        
        '---------------------------------------------------------------------------
        
        itmX.SubItems(3) = Format(RsAux!CDiImporte, FormatoMonedaP)
        aTotal = aTotal + RsAux!CDiImporte
        
        itmX.SubItems(4) = Format(RsAux!CDiLibrado, "dd/mm/yy")
        
        If oADepositar.Value Then
            'Sucursal a Depositar SBaDeposito-----------------------------------------------------------
            If aSucCodigo <> RsAux!SBaDeposito Then
                Cons = "Select * from  SucursalDeBanco, BancoSSFF" _
                        & " Where SBaCodigo = " & RsAux!SBaDeposito _
                        & " And SBaBanco = BanCodigo"
                Set RsSuc = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                aSucCodigo = RsAux!SBaDeposito
                aSucNombre = Trim(RsSuc!BanNombre) & " (" & Trim(RsSuc!SBaNombre) & ")"
                RsSuc.Close
            End If
            '-------------------------------------------------------------------------------------------------
        End If
        If oDepositado.Value Then
            'Sucursal Depositado CDiDepositado-----------------------------------------------------------
            If aSucCodigo <> RsAux!CDiDepositado Then
                Cons = "Select * from  SucursalDeBanco, BancoSSFF" _
                        & " Where SBaCodigo = " & RsAux!CDiDepositado _
                        & " And SBaBanco = BanCodigo"
                Set RsSuc = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                aSucCodigo = RsAux!CDiDepositado
                aSucNombre = Trim(RsSuc!BanNombre) & " (" & Trim(RsSuc!SBaNombre) & ")"
                RsSuc.Close
            End If
            '-------------------------------------------------------------------------------------------------
        End If
        
        itmX.Tag = Trim(itmX.Tag) & "S" & aSucCodigo
        itmX.SubItems(5) = aSucNombre
        
        If bConsultar.Tag = "DE" Then itmX.SubItems(6) = Format(RsAux!CDiCobrado, "dd/mm/yy")
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    lTotal.Caption = Format(aTotal, FormatoMonedaP)
    If lLista.ListItems.Count > 0 And bConsultar.Tag = "AD" Then bGrabar.Enabled = True
    Screen.MousePointer = 0
    Exit Sub
    
errPago:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los cheques diferidos."
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    If Me.Width >= 9120 Then
        
        lTTotal.Left = Me.Width - lTTotal.Width - 220
        lTotal.Left = Me.Width - lTotal.Width - 250
        
        lTTotal.Top = Me.Height - lTTotal.Height - 460
        lTotal.Top = Me.Height - lTotal.Height - 460
        
        lLista.Width = Me.Width - 350
        lLista.Height = Me.Height - lLista.Top - lTotal.Height - 550
        Cuadro.Width = Me.Width - 350
                
        bGrabar.Top = lTotal.Top - 60
        bGrabar.Left = lTTotal.Left - bGrabar.Width - 100
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Forms(Forms.Count - 2).SetFocus
End Sub

Private Sub Label1_Click()
    Foco tFecha
End Sub

Private Sub Label2_Click()
    Foco cMoneda
End Sub

Private Sub lLista_KeyDown(KeyCode As Integer, Shift As Integer)

    If lLista.ListItems.Count = 0 Then Exit Sub
    On Error GoTo errLista
    If KeyCode = vbKeyDelete And bConsultar.Tag = "DE" Then
        'Elimino depósito del cheque---------------------------------------------------
        Set itmX = lLista.SelectedItem
        If MsgBox("Confirma eliminar el depósito del cheque: " & lLista.SelectedItem.SubItems(2), vbQuestion + vbYesNo, "ELIMINAR") = vbNo Then Exit Sub
        Screen.MousePointer = 11
        Cons = "Update ChequeDiferido " _
                & " Set CDiCobrado = NULL , CDiDepositado = NULL" _
                & " Where CDiCodigo = " & Mid(itmX.Key, 2, Len(itmX.Key))
        cBase.Execute Cons
        lLista.ListItems.Remove lLista.SelectedItem.Index
        Screen.MousePointer = 0
    End If      '--------------------------------------------------------------------------------------------
    Exit Sub
    
errLista:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al procesar la información."
End Sub

Private Sub lLista_KeyPress(KeyAscii As Integer)
    
    If lLista.ListItems.Count = 0 Then Exit Sub
    On Error GoTo errLista
    Select Case KeyAscii
        Case vbKeySpace: CambioIcono
        Case vbKeyReturn: If bGrabar.Enabled Then bGrabar.SetFocus
    End Select
    
    Exit Sub
errLista:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al procesar la información."
End Sub

Private Sub lLista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    Set lLista.SelectedItem = lLista.HitTest(X, Y)
    If lLista.ListItems.Count = 0 Then Exit Sub
    If bConsultar.Tag = "DE" Then MnuVerSucursal.Enabled = False Else: MnuVerSucursal.Enabled = True
    If Button = vbRightButton Then PopupMenu MnuBDerecho

End Sub

Private Sub MnuVerCheque_Click()
    
    On Error GoTo errCheque
    Screen.MousePointer = 11
    Dim frmCheque As New CoChequeDiferido
    
    Set itmX = lLista.SelectedItem
    frmCheque.pSeleccionado = CLng(Mid(itmX.Key, 2, Len(itmX.Key)))
    frmCheque.Show vbModal, Me
    
    Set frmCheque = Nothing
    Screen.MousePointer = 0
    Exit Sub
    
errCheque:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al visualizar el cheque."
End Sub

Private Sub MnuVerDeuda_Click()
    EjecutarApp pathApp & "\Deuda en cheques", ClienteDelTag(lLista.SelectedItem.Tag)
End Sub

Private Sub MnuVerSucursal_Click()
    CambioBancoADepositar
End Sub

Private Sub oADepositar_Click()
    
    If oADepositar.Value Then lTTotal.Caption = " Total a Depositar:"
            
End Sub

Private Sub oADepositar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub oADepositar_LostFocus()
    If oADepositar.Value Then lTTotal.Caption = " Total a Depositar:" Else lTTotal.Caption = " Total Depositado:"
End Sub

Private Sub oDepositado_Click()
    If oDepositado.Value Then lTTotal.Caption = " Total Depositado:"
End Sub

Private Sub oDepositado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub oDepositado_LostFocus()
    If oDepositado.Value Then lTTotal.Caption = " Total Depositado:" Else lTTotal.Caption = " Total a Depositar:"
End Sub

Private Sub tFecha_GotFocus()
    tFecha.SelStart = 0
    tFecha.SelLength = Len(tFecha.Text)
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        aTexto = ValidoPeriodoFechas(tFecha.Text)
        If aTexto = "" Then MsgBox "El período de fechas ingresado no es correcto.", vbExclamation, "ATENCIÓN": Exit Sub
        Foco cMoneda
        If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, "d/mm/yyyy")
    End If
    
End Sub

Private Sub CambioIcono()
    
    On Error Resume Next
    If lLista.SelectedItem.Index = -1 Then Exit Sub
    MarcoRedraw 0, lLista
    Set itmX = lLista.SelectedItem
    
    If itmX.SmallIcon = "check" Then
        itmX.SmallIcon = "nocheck"
        lTotal.Caption = Format(CCur(lTotal.Caption) - CCur(itmX.SubItems(3)), FormatoMonedaP)
    Else
        itmX.SmallIcon = "check"
        lTotal.Caption = Format(CCur(lTotal.Caption) + CCur(itmX.SubItems(3)), FormatoMonedaP)
    End If
    
    MarcoRedraw 1, lLista
    
End Sub

Private Sub CambioBancoADepositar()
    
    On Error Resume Next
    If lLista.SelectedItem.Index = -1 Then Exit Sub
    
    Dim aCodigo As String
    Dim aCodSucursal As Long
    
    aCodigo = InputBox("Ingrese el codigo de banco y sucursal en donde se va a depositar el cheque (Formato 00-000).", "Sucursal a Depositar")
    
    If Trim(aCodigo) = "" Then Exit Sub
    If InStr(aCodigo, "-") = 0 Then MsgBox "El código ingresado no es correcto.", vbExclamation, "ATENCIÓN": Exit Sub
    
    On Error GoTo errCargar
    Dim aSucursal As Integer, aBanco As Integer
    aBanco = CInt(Mid(aCodigo, 1, InStr(aCodigo, "-") - 1))
    aSucursal = CInt(Mid(aCodigo, InStr(aCodigo, "-") + 1, Len(aCodigo)))
    
    Screen.MousePointer = 11
    Cons = "Select * from  SucursalDeBanco, BancoSSFF" _
            & " Where SBaCodigoS = " & aSucursal _
            & " And BanCodigoB = " & aBanco _
            & " And SBaBanco = BanCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        aCodigo = Trim(RsAux!BanNombre) & " (" & Trim(RsAux!SBaNombre) & ")"
        aCodSucursal = RsAux!SBaCodigo
    Else
        aCodigo = ""
        Screen.MousePointer = 0
        MsgBox "No existe registro para el código ingresado.", vbExclamation, "ATENCIÓN"
    End If
    RsAux.Close
    
    If aCodigo <> "" Then
        MarcoRedraw 0, lLista
        Set itmX = lLista.SelectedItem
        itmX.SubItems(5) = aCodigo
        itmX.Tag = Mid(itmX.Tag, 1, InStr(itmX.Tag, "S")) & aCodSucursal
        MarcoRedraw 1, lLista
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al buscar el banco."
End Sub

Private Sub AccionGrabar()

    FechaDelServidor
    
    On Error GoTo errorBT
    Screen.MousePointer = 11
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
    
    For Each itmX In lLista.ListItems
        If itmX.SmallIcon = "check" Then
            Cons = "Update ChequeDiferido " _
                    & " Set CDiCobrado = '" & Format(gFechaServidor, FormatoFH) & "', " _
                    & " CDiDepositado = " & SucursalDelTag(itmX.Tag) _
                    & " Where CDiCodigo = " & Mid(itmX.Key, 2, Len(itmX.Key))
            cBase.Execute Cons
        End If
    Next
    
    cBase.CommitTrans    'Fin de la TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    bGrabar.Enabled = False
    lLista.ListItems.Clear
    Screen.MousePointer = 0
    
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

Private Function SucursalDelTag(aTag As String) As Long
    SucursalDelTag = CLng(Trim(Mid(aTag, InStr(aTag, "S") + 1, Len(aTag))))
End Function
Private Function ClienteDelTag(aTag As String) As Long
    ClienteDelTag = CLng(Trim(Mid(aTag, InStr(aTag, "C") + 1, InStr(aTag, "S") - InStr(aTag, "C") - 1)))
End Function

