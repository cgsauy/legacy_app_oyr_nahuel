VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compra de Moneda Extranjera"
   ClientHeight    =   3630
   ClientLeft      =   4575
   ClientTop       =   3195
   ClientWidth     =   3435
   ClipControls    =   0   'False
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   3435
   Begin VB.TextBox tDoc 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   60
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   3300
      Width           =   1095
   End
   Begin VB.CommandButton bGrabar 
      Caption         =   "&Grabar"
      Height          =   315
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3300
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6015
      TabIndex        =   15
      Top             =   -60
      Width           =   6015
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3000
         Picture         =   "frmMain.frx":030A
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   21
         Top             =   120
         Width           =   300
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   45
         Picture         =   "frmMain.frx":05C5
         Stretch         =   -1  'True
         Top             =   90
         Width           =   420
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label8"
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   0
         TabIndex        =   17
         Top             =   480
         Width           =   6795
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " Compra de M/E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   480
         TabIndex        =   16
         Top             =   120
         Width           =   2955
      End
   End
   Begin VB.ComboBox cSignos 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3960
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox tImporteCotizado 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   2340
      Width           =   1095
   End
   Begin VB.ComboBox cMonedaD 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   900
      Width           =   2235
   End
   Begin VB.TextBox tImporteE 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1140
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2340
      Width           =   1035
   End
   Begin VB.TextBox tCotizacion 
      Height          =   315
      Left            =   1080
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1260
      Width           =   975
   End
   Begin VB.TextBox tImporteD 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1980
      Width           =   1095
   End
   Begin VB.ComboBox cMonedaE 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   540
      Width           =   2235
   End
   Begin VB.Label lPie 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   15
      Left            =   480
      TabIndex        =   19
      Top             =   3180
      Width           =   2295
   End
   Begin VB.Label lSubTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   60
      TabIndex        =   14
      Top             =   2760
      Width           =   3315
   End
   Begin VB.Label lMD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Para Pagar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2340
      TabIndex        =   12
      Top             =   1680
      Width           =   1035
   End
   Begin VB.Label lME 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cambia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1140
      TabIndex        =   11
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Se &Venden"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   960
      Width           =   915
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Se &Compran"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cam&bia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   2400
      Width           =   915
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cotizados &A"
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Para Cobrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.Menu MnuAcciones 
      Caption         =   "MnuAcciones"
      Visible         =   0   'False
      Begin VB.Menu MnuTitulo 
         Caption         =   "MnuTitulo"
      End
      Begin VB.Menu MnuAL0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPendiente 
         Caption         =   "Dejar como &Pendiente"
      End
      Begin VB.Menu MnuMExtranjera 
         Caption         =   "Tomar Moneda &Extranjera"
      End
      Begin VB.Menu MnuAL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuChequesDif 
         Caption         =   "Ingresar Cheques &Diferidos"
      End
      Begin VB.Menu MnuChequesAlD 
         Caption         =   "Ingresar Cheques Al &Día"
      End
      Begin VB.Menu MnuAL2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAnulaciones 
         Caption         =   "&Anulación de Documentos"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public prmIDDocumento As Long
Public prmIDMonedaV As Integer
Public prmImporteV As Currency
Public prmTC As Currency
Public prmComs As String

Dim mTexto As String
Dim oCnfgPrint As New clsImpresoraTicketsCnfg

Private bGrabando As Boolean    'Por doble enter

'Variables para Crystal Engine.---------------------------------
Private result As Integer, JobSRep1 As Integer, JobSRep2 As Integer, jobnum As Integer
Private NombreFormula As String, CantForm As Integer, aTexto As String

Private Sub bGrabar_Click()
    AccionGrabar
End Sub

Private Sub cMonedaD_Click()
    tCotizacion.Text = ""
    lMD.Caption = "..."
    tImporteCotizado.Text = ""
    lSubTotal.Caption = ""
End Sub

Private Sub cMonedaD_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cMonedaD.ListIndex <> -1 Then
            CargoValoresMoneda
            Foco tImporteE
        End If
    End If
End Sub

Private Sub cMonedaE_Click()
    tCotizacion.Text = ""
    lME.Caption = "..."
    tImporteE.Text = ""
    tImporteCotizado.Text = ""
    lSubTotal.Caption = ""
End Sub

Private Sub cMonedaE_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cMonedaE.ListIndex <> -1 Then
            CargoValoresMoneda
            Foco tImporteE
        End If
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub

Private Sub Form_Load()
On Error GoTo errLoad

    bGrabando = False
    LimpioTodo
    FechaDelServidor
    dis_CargoArrayMonedas
    
    
    oCnfgPrint.CargarConfiguracion prmKeyApp
    
    Me.BackColor = RGB(176, 196, 222)
    ObtengoSeteoForm Me
    Me.Height = 4035: Me.Width = 3555
    
    InicializoControles
    
    If prmIDDocumento <> 0 Or prmIDMonedaV <> 0 Or prmImporteV <> 0 Then
        CargoValoresIngreso
    Else
        BuscoCodigoEnCombo cMonedaD, CLng(paMonedaPesos)
        If cMonedaD.ListIndex <> -1 Then CargoValoresMoneda
    End If
    
    Exit Sub
errLoad:
    clsGeneral.OcurrioError "Error al iniciar el formulario.", Err.Description
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    lPie.Left = Me.ScaleLeft: lPie.Width = Me.ScaleWidth
    
    tDoc.Width = bGrabar.Left - tDoc.Left - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    GuardoSeteoForm Me
    EndMain
End Sub


Private Sub InicializoControles()
    
    On Error Resume Next
    Cons = "Select MonCodigo, MonNombre, MonSigno from Moneda Order by MonNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    Do While Not RsAux.EOF
        With cMonedaE
            .AddItem Trim(RsAux!MonNombre)
            .ItemData(.NewIndex) = RsAux!MonCodigo
        End With
        With cMonedaD
            .AddItem Trim(RsAux!MonNombre)
            .ItemData(.NewIndex) = RsAux!MonCodigo
        End With
        
        With cSignos
            .AddItem Trim(RsAux!MonSigno)
            .ItemData(.NewIndex) = RsAux!MonCodigo
        End With
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Dim bkColor As Long
    bkColor = RGB(240, 255, 255)
    cMonedaD.BackColor = bkColor
    cMonedaE.BackColor = bkColor
    tImporteD.BackColor = bkColor: tCotizacion.BackColor = bkColor
    tImporteE.BackColor = bkColor
    tImporteCotizado.BackColor = bkColor
    lSubTotal.BackColor = bkColor
    
    bGrabar.BackColor = Me.BackColor
    tDoc.BackColor = RGB(180, 217, 240)
End Sub

Private Sub CargoValoresIngreso()

Dim mMoneda As Long, mTotal As Currency

    If prmIDDocumento <> 0 Then
        Cons = "Select * from Documento Where DocCodigo = " & prmIDDocumento
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            mMoneda = RsAux!DocMoneda
            mTotal = RsAux!DocTotal
            tDoc.Text = Trim(RsAux!DocSerie) & "-" & RsAux!DocNumero
            tDoc.Tag = RsAux!DocCodigo
        End If
        RsAux.Close
    
        If mMoneda = 0 Then Exit Sub
    
        tImporteD.Text = Format(mTotal, "#,##0.00")
        BuscoCodigoEnCombo cMonedaD, mMoneda
    
    Else
        Dim bInvertir As Boolean
        If prmIDMonedaV = paMonedaPesos Then bInvertir = True
        
        mMoneda = paMonedaPesos
        mTotal = prmImporteV
        If prmIDMonedaV <> 0 Then
            'BuscoCodigoEnCombo cMonedaE, CLng(prmIDMonedaV)
            If Not bInvertir Then
                BuscoCodigoEnCombo cMonedaE, CLng(prmIDMonedaV)
            Else
                BuscoCodigoEnCombo cMonedaE, CLng(paMonedaDolar)
            End If
        End If
        
        BuscoCodigoEnCombo cMonedaD, mMoneda
        
        If Not bInvertir Then
            tImporteE.Text = Format(mTotal, "#,##0.00")
        Else
            tImporteD.Text = Format(mTotal, "#,##0.00")
        End If
        
'        tImporteE.Text = Format(mTotal, "#,##0.00")
        
        If prmTC <> 0 Then
            mTotal = mTotal * prmTC
            tCotizacion.Text = prmTC
            tImporteD.Text = Format(mTotal, "#,##0.00")
        End If
        
        tDoc.Text = Trim(prmComs)
    End If
    
    
    If cMonedaD.ListIndex <> -1 Then CargoValoresMoneda
    
End Sub

Private Sub CargoValoresMoneda()
On Error Resume Next
    Dim mIDMoneda As Integer, mIDMonedaE As Integer
    
    mIDMoneda = cMonedaD.ItemData(cMonedaD.ListIndex)
    
    If cMonedaE.ListIndex = -1 Then
        BuscoCodigoEnCombo cMonedaE, CLng(paMonedaDolar)
    End If
    
    If cMonedaE.ListIndex <> -1 Then mIDMonedaE = cMonedaE.ItemData(cMonedaE.ListIndex)
    
    Dim mValor As Currency
    'Busco la última tasa de cambio
    If cMonedaE.ListIndex <> -1 Then
        If Not IsNumeric(tCotizacion.Text) Then
            mValor = TasadeCambio(mIDMonedaE, mIDMoneda, gFechaServidor, TipoTC:=prmTipoTC)
        Else
            mValor = CCur(tCotizacion.Text)
        End If
        tCotizacion.Text = Format(mValor, "#,##0.000")
    End If
    
    If Trim(tImporteD.Text) = "" And IsNumeric(tImporteE.Text) And IsNumeric(tCotizacion.Text) Then
        tImporteD.Text = CCur(tImporteE.Text) * CCur(tCotizacion.Text)
    End If
    tImporteD.Text = Format(tImporteD.Text, "#,##0.00")
    
    
    If IsNumeric(tCotizacion.Text) And (IsNumeric(tImporteD.Text) Or IsNumeric(tImporteE.Text)) Then
        
        If Not IsNumeric(tImporteE.Text) Then
            mValor = CCur(tImporteD.Text) / CCur(tCotizacion.Text)
        Else
            mValor = tImporteE.Text
        End If
        tImporteE.Text = Redondeo(mValor, dis_arrMonedaProp(CLng(mIDMonedaE), pRedondeo))
        tImporteE.Text = Format(tImporteE.Text, "#,##0.00")
        
        mValor = CCur(tImporteE.Text) * CCur(tCotizacion.Text)
        tImporteCotizado.Text = Redondeo(mValor, dis_arrMonedaProp(CLng(mIDMoneda), pRedondeo))
        tImporteCotizado.Text = Format(tImporteCotizado.Text, "#,##0.00")
    
    End If
    
    If IsNumeric(tImporteCotizado.Text) And IsNumeric(tImporteD.Text) Then
        mValor = CCur(tImporteCotizado.Text) - CCur(tImporteD.Text)
        lSubTotal.Tag = mValor
        
        If mValor > 0 Then
            lSubTotal.Caption = "Devolver "
            lSubTotal.ForeColor = &H8000&
            lSubTotal.FontBold = False
        Else
            lSubTotal.Caption = "Reclamar "
            lSubTotal.ForeColor = Colores.RojoClaro
            lSubTotal.FontBold = True
        End If
        If Abs(mValor) <> 0 Then
            mTexto = Format(Abs(mValor), "#,##0.00")
            If Right(mTexto, 3) = ".00" Then mTexto = Mid(mTexto, 1, Len(mTexto) - 3)
            lSubTotal.Caption = lSubTotal.Caption & mTexto
        Else
            lSubTotal.Caption = ""
        End If
    End If
    
    
    If cMonedaE.ListIndex <> -1 Then
        BuscoCodigoEnCombo cSignos, cMonedaE.ItemData(cMonedaE.ListIndex)
        lME.Caption = Trim(cSignos.Text)
    End If
    If cMonedaD.ListIndex <> -1 Then
        BuscoCodigoEnCombo cSignos, cMonedaD.ItemData(cMonedaD.ListIndex)
        lMD.Caption = Trim(cSignos.Text)
    End If
       
End Sub

Private Sub LimpioTodo()

    cMonedaD.ListIndex = -1
    tImporteD.Text = "": tCotizacion.Text = ""
    cMonedaE.ListIndex = -1
    tImporteE.Text = ""
    tImporteCotizado.Text = ""
    lSubTotal.Caption = ""
    tDoc.Text = ""
    
End Sub

Private Sub Picture2_Click()
    frmDondeImprimo.Show vbModal
    oCnfgPrint.CargarConfiguracion prmKeyApp
    Picture2.Refresh
    Me.Refresh
End Sub

Private Sub tCotizacion_Change()
    If Trim(tImporteCotizado.Text) <> "" Then tImporteCotizado.Text = ""
End Sub

Private Sub tCotizacion_GotFocus()
    tCotizacion.SelStart = 0: tCotizacion.SelLength = Len(tCotizacion.Text)
End Sub

Private Sub tCotizacion_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        CargoValoresMoneda
        Foco tImporteE
    End If
    
End Sub

Private Sub tImporteD_Change()
    If Trim(tImporteCotizado.Text) <> "" Then tImporteCotizado.Text = ""
End Sub

Private Sub tImporteD_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        CargoValoresMoneda
        AccionGrabar
    End If
End Sub

Private Sub tImporteE_Change()
    If Trim(tImporteCotizado.Text) <> "" Then tImporteCotizado.Text = ""
End Sub

Private Sub tImporteE_GotFocus()
    tImporteE.SelStart = 0: tImporteE.SelLength = Len(tImporteE.Text)
End Sub

Private Sub tImporteE_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        CargoValoresMoneda
        Foco tImporteD
    End If
    
End Sub

Private Sub AccionGrabar()
On Error GoTo errorComun

    If bGrabando Then Exit Sub
    'Si compro Dolares: Salen Pesos Entran Dolares
    If Not ValidoDatos Then Exit Sub
    
    If Trim(tDoc.Text) <> "" Then mTexto = "Docs: " & Trim(tDoc.Text) & "; "
    mTexto = "Compra ME " & mTexto & "Usr: " & miConexion.UsuarioLogueado(Nombre:=True)
    
    Dim sImporte As String
    If Trim(lSubTotal.Caption) <> "" And Val(lSubTotal.Tag) <> 0 Then
        If Val(lSubTotal.Tag) > 0 Then
            sImporte = "Se Devuelve "
        Else
            sImporte = "Se Cobra "
        End If
        
        sImporte = sImporte & "  " & lMD.Caption & " " & Format(Abs(lSubTotal.Tag), "#,##0.00")
        
        mTexto = mTexto & " [[" & sImporte & "]]"
    End If
    
    Dim TC As Currency
    Dim aMovimiento As Long
    Dim mDSalida As Long, mDEntrada As Long
    Dim mMonedaS As Long, mMonedaE As Long
    Dim mImporteS As Currency, mImporteE As Currency
    Dim mImportePesos As Currency
    
    mMonedaE = cMonedaE.ItemData(cMonedaE.ListIndex)
    mMonedaS = cMonedaD.ItemData(cMonedaD.ListIndex)
    
    mImporteS = CCur(tImporteCotizado.Text)
    mImporteE = CCur(tImporteE.Text)
    
    mDEntrada = dis_DisponibilidadPara(paCodigoDeSucursal, mMonedaE)
    mDSalida = dis_DisponibilidadPara(paCodigoDeSucursal, mMonedaS)
    
    If mDEntrada = 0 Then
        MsgBox "No se puede realizar la compra de Moneda Extranjera." & vbCrLf & _
                    "No existe una disponibilidad para dar el ingreso de '" & cMonedaE.Text & "'.", vbInformation, "Falta Disponibilidad"
        Exit Sub
    End If
    
    If mDSalida = 0 Then
        MsgBox "No se puede realizar la compra de Moneda Extranjera." & vbCrLf & _
                    "No existe una disponibilidad para dar la salida de '" & cMonedaD.Text & "'.", vbInformation, "Falta Disponibilidad"
        Exit Sub
    End If
    
    If MsgBox("Confirma realizar la compra de Moneda Extranjera.", vbQuestion + vbYesNo, "Grabar Compra") = vbNo Then Exit Sub
    
    bGrabando = True
    Screen.MousePointer = 11
    
    'Si una de las dos monedas es pesos cargo el valor importe en pesos    --------------------------
    mImportePesos = 0
    If mMonedaS = paMonedaPesos Then
        mImportePesos = mImporteS
    ElseIf mMonedaE = paMonedaPesos Then
        mImportePesos = mImporteE
    End If
    '------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errorBT
    cBase.BeginTrans            'COMIENZO TRANSACCION------------------------------------------     !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
    
    'Inserto en la Tabla Movimiento-Disponibilidad--------------------------------------------------------
    Cons = "Select * from MovimientoDisponibilidad Where MDiID = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.AddNew
    RsAux!MDiFecha = Format(gFechaServidor, "mm/dd/yyyy")
    RsAux!MDiHora = Format(gFechaServidor, "hh:mm:ss")
    RsAux!MDiTipo = prmMCCompraME
    RsAux!MDiComentario = Trim(mTexto)
    RsAux.Update: RsAux.Close
    '------------------------------------------------------------------------------------------------------------
    
    'Saco el Id de movimiento-------------------------------------------------------------------------------
    Cons = "Select Max(MDiID) from MovimientoDisponibilidad" & _
               " Where MDiTipo = " & prmMCCompraME
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    aMovimiento = RsAux(0)
    RsAux.Close
    '------------------------------------------------------------------------------------------------------------
        
    'Grabo en Tabla Movimiento-Disponibilidad-Renglon--------------------------------------------------
    Cons = "Select * from MovimientoDisponibilidadRenglon Where MDRIdMovimiento = " & aMovimiento
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    '1) Grabo la SALIDA    -------------------------------
    RsAux.AddNew
    RsAux!MDRIdMovimiento = aMovimiento
    RsAux!MDRIdDisponibilidad = mDSalida
    RsAux!MDRIdCheque = 0
    
    RsAux!MDRImporteCompra = mImporteS
    RsAux!MDRHaber = mImporteS
    
    If mImportePesos <> 0 Then
        RsAux!MDRImportePesos = mImportePesos
    Else
        TC = TasadeCambio(CInt(mMonedaS), paMonedaPesos, gFechaServidor)           'Tasa de cambio a pesos
        RsAux!MDRImportePesos = mImporteS * TC
    End If
    RsAux.Update
    
    '2) Grabo la ENTRADA    -------------------------------
    RsAux.AddNew
    RsAux!MDRIdMovimiento = aMovimiento
    RsAux!MDRIdDisponibilidad = mDEntrada
    RsAux!MDRIdCheque = 0
    
    RsAux!MDRImporteCompra = mImporteE
    RsAux!MDRDebe = mImporteE
    
    If mImportePesos <> 0 Then
        RsAux!MDRImportePesos = mImportePesos
    Else
        TC = TasadeCambio(CInt(mMonedaE), paMonedaPesos, gFechaServidor)          'Tasa de cambio a pesos
        RsAux!MDRImportePesos = mImporteE * TC
    End If
    
    RsAux.Update
    
    RsAux.Close
    '------------------------------------------------------------------------------------------------------------
    
    If oCnfgPrint.Opcion > 0 And aMovimiento > 0 Then
        '[dbo].[prg_PosInsertoDocumentosATickets] @documentos varchar(5000), @impresora tinyint, @importe money = null
        cBase.Execute "EXEC prg_PosInsertoDocumentosATickets '" & aMovimiento * -1 & "', " & oCnfgPrint.ImpresoraTickets & ", " & IIf(IsNumeric(tImporteD), CCur(tImporteD.Text), 0)
    End If
    
    cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------        !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    If oCnfgPrint.Opcion = 0 Then AccionImprimir
    
    Unload Me
    Screen.MousePointer = 0
    Exit Sub

errorComun:
    clsGeneral.OcurrioError "Error.", Err.Description
    Screen.MousePointer = 0: bGrabando = False
    Exit Sub
errorBT:
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
    Screen.MousePointer = 0: bGrabando = False: Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Error al grabar los datos.", Err.Description
    Screen.MousePointer = 0: bGrabando = False
End Sub

Private Function ValidoDatos() As Boolean
    ValidoDatos = False
    
    If cMonedaD.ListIndex = -1 Then
        MsgBox "Debe seleccionar la moneda del cambio.", vbExclamation, "Faltan Datos"
        cMonedaD.SetFocus: Exit Function
    End If
    If cMonedaE.ListIndex = -1 Then
        MsgBox "Debe seleccionar la moneda a comprar.", vbExclamation, "Faltan Datos"
        cMonedaE.SetFocus: Exit Function
    End If
    
    If cMonedaE.ItemData(cMonedaE.ListIndex) = cMonedaD.ItemData(cMonedaD.ListIndex) Then
        MsgBox "Las monedas para realizar la compra/venta no deben ser iguales.", vbExclamation, "Posible Error"
        cMonedaE.SetFocus: Exit Function
    End If
    
    If Not IsNumeric(tCotizacion.Text) Then
        MsgBox "La cotización ingresada no es correcta.", vbExclamation, "Falta Cotización"
        tCotizacion.SetFocus: Exit Function
    End If
    
    If Not IsNumeric(tImporteE.Text) Then
        MsgBox "El importe en " & lME.Caption & " no es correcto.", vbExclamation, "Faltan Datos"
        tImporteE.SetFocus: Exit Function
    End If
     If Not IsNumeric(tImporteD.Text) Then
        MsgBox "El importe en " & lMD.Caption & " no es correcto.", vbExclamation, "Faltan Datos"
        tImporteD.SetFocus: Exit Function
    End If
    If CCur(tImporteE.Text) = 0 Then
        MsgBox "El importe en " & lME.Caption & " no es correcto.", vbExclamation, "Faltan Datos"
        tImporteE.SetFocus: Exit Function
    End If
     If CCur(tImporteD.Text) = 0 Then
        MsgBox "El importe en " & lMD.Caption & " no es correcto.", vbExclamation, "Faltan Datos"
        tImporteD.SetFocus: Exit Function
    End If
    
    If Not IsNumeric(tImporteCotizado.Text) Then
        MsgBox "Faltan ingresar datos para establecer el importe cotizado.", vbExclamation, "Psible Error"
        tImporteE.SetFocus: Exit Function
    End If
    
    ValidoDatos = True
End Function


Private Sub AccionImprimir()

On Error GoTo errCrystal

    Screen.MousePointer = 11
    Dim bOK As Boolean
    crAbroEngine
    bOK = ImprimoReciboCompra
    crCierroEngine
    If Not bOK Then Resume errCrystal
    
    Screen.MousePointer = 0
    Exit Sub

errCrystal:
    clsGeneral.OcurrioError crMsgErr
    Screen.MousePointer = 0
End Sub

Private Function ImprimoReciboCompra() As Boolean
On Error GoTo errPrint

    ImprimoReciboCompra = False
    'jobnum = crAbroReporte(prmPathListados & "Recibo.RPT")
    jobnum = crAbroReporte(prmPathListados & "blankRecibo.RPT")
    If jobnum = 0 Then GoTo errPrint
    
    If ChangeCnfgPrint Then prj_LoadConfigPrint bShowFrm:=False
    
    'Configuro la Impresora
    If Trim(Printer.DeviceName) <> Trim(paIReciboN) Then SeteoImpresoraPorDefecto paIReciboN
    'If Not crSeteoParaRecibos(jobnum, Printer, paIReciboB) Then GoTo errPrint
    If Not crSeteoImpresora(jobnum, Printer, paIReciboB, paperSize:=13, mOrientation:=2) Then GoTo errPrint
    
    
    'Obtengo la cantidad de formulas que tiene el reporte.----------------------
    CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
    If CantForm = -1 Then GoTo errPrint
    
    'Cargo Propiedades para el reporte Contado --------------------------------
    For I = 0 To CantForm - 1
        NombreFormula = crObtengoNombreFormula(jobnum, CInt(I))
        
        Select Case LCase(NombreFormula)
            Case "": GoTo errPrint
            Case "nombredocumento": result = crSeteoFormula(jobnum%, NombreFormula, "''")
                
            Case "docs": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(tDoc.Text) & "'")
            Case "totaldocs"
                                mTexto = Trim(lMD.Caption) & " " & tImporteD.Text
                                result = crSeteoFormula(jobnum%, NombreFormula, "'" & mTexto & "'")
            
            Case "totalme"
                                mTexto = Trim(lME.Caption) & " " & tImporteE.Text
                                result = crSeteoFormula(jobnum%, NombreFormula, "'" & mTexto & "'")
                        
             Case "cotizadosa"
                                mTexto = " cotizados a " & Trim(lMD.Caption) & " " & tCotizacion.Text & " = " & Trim(lMD.Caption) & " " & tImporteCotizado.Text
                                result = crSeteoFormula(jobnum%, NombreFormula, "'" & mTexto & "'")
                                     
            Case "texto1"
                                mTexto = ""
                                If Trim(lSubTotal.Caption) <> "" And Val(lSubTotal.Tag) <> 0 Then
                                    If Val(lSubTotal.Tag) > 0 Then
                                        mTexto = "Se Devuelve "
                                    Else
                                        mTexto = "Se Cobra "
                                    End If
                                End If
                                result = crSeteoFormula(jobnum%, NombreFormula, "'" & mTexto & "'")
                                
            Case "texto2"
                                mTexto = ""
                                If Trim(lSubTotal.Caption) <> "" And Val(lSubTotal.Tag) <> 0 Then
                                    mTexto = Format(Abs(lSubTotal.Tag), "#,##0.00")
                                    mTexto = lMD.Caption & " " & mTexto
                                End If
                                result = crSeteoFormula(jobnum%, NombreFormula, "'" & mTexto & "'")
                                
            
            Case "fecha":
                                mTexto = Format(gFechaServidor, "dd/mm/yy")
                                result = crSeteoFormula(jobnum%, NombreFormula, "'" & mTexto & "'")
            
            
            
            Case Else: result = 1
        End Select
        If result = 0 Then GoTo errPrint
    Next
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    'Seteo la Query del reporte-----------------------------------------------------------------
    Cons = "SELECT Documento.DocCodigo , Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal, Documento.DocIVA, Documento.DocVendedor" _
            & " From CGSA.dbo.Documento Documento " _
            & " Where DocCodigo = " & 0
    If crSeteoSqlQuery(jobnum%, Cons) = 0 Then GoTo errPrint
        
    'If crMandoAPantalla(jobnum, "Recibo Compra ME") = 0 Then GoTo errPrint
    If crMandoAImpresora(jobnum, 1) = 0 Then GoTo errPrint
    If Not crInicioImpresion(jobnum, True, False) Then GoTo errPrint
    
    'crEsperoCierreReportePantalla
    
    If Not crCierroTrabajo(jobnum) Then MsgBox crMsgErr
    
    ImprimoReciboCompra = True
    Exit Function

errPrint:
    On Error Resume Next
    crCierroTrabajo (jobnum)
End Function
