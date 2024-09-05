VERSION 5.00
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form InTasaCambio 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TC - Tipos de Cambio"
   ClientHeight    =   5145
   ClientLeft      =   4635
   ClientTop       =   2340
   ClientWidth     =   6030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "InTasaCambio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6030
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6015
      TabIndex        =   16
      Top             =   0
      Width           =   6015
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " Ingreso de Cotizaciones"
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
         TabIndex        =   18
         Top             =   120
         Width           =   2955
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
      Begin VB.Image Image1 
         Height          =   420
         Left            =   40
         Picture         =   "InTasaCambio.frx":030A
         Stretch         =   -1  'True
         Top             =   30
         Width           =   420
      End
   End
   Begin MSComCtl2.DTPicker dBuscar 
      Height          =   315
      Left            =   2760
      TabIndex        =   13
      Top             =   780
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   23658497
      CurrentDate     =   37481
   End
   Begin VB.CommandButton bExit 
      Caption         =   "&Salir"
      Height          =   315
      Left            =   4920
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   780
      Width           =   1035
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
      Height          =   3915
      Left            =   2760
      TabIndex        =   14
      Top             =   1140
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   6906
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
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   5
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
   Begin VB.TextBox tVendedor 
      Height          =   315
      Left            =   1440
      TabIndex        =   11
      Top             =   3480
      Width           =   1155
   End
   Begin VB.TextBox tComprador 
      Height          =   315
      Left            =   60
      TabIndex        =   9
      Top             =   3480
      Width           =   1155
   End
   Begin VB.TextBox tFecha 
      Height          =   315
      Left            =   60
      MaxLength       =   12
      TabIndex        =   7
      Top             =   2880
      Width           =   1155
   End
   Begin AACombo99.AACombo cOrigen 
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   1440
      Width           =   2295
      _ExtentX        =   4048
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
      Text            =   ""
   End
   Begin AACombo99.AACombo cDestino 
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   2100
      Width           =   2295
      _ExtentX        =   4048
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
      Text            =   ""
   End
   Begin AACombo99.AACombo cCotizacion 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   780
      Width           =   2535
      _ExtentX        =   4471
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
      Text            =   ""
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "&Buscar Cotizaciones"
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   540
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "&Tipo de Cotización"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Venta"
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Co&mpra"
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Fecha"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Cotizados En"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   1860
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Moneda"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
End
Attribute VB_Name = "InTasaCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bExit_Click()
    Unload Me
End Sub

Private Sub cCotizacion_Click()
    vsLista.Rows = 1
End Sub

Private Sub cCotizacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cOrigen
End Sub

Private Sub cDestino_Click()
    vsLista.Rows = 1
End Sub

Private Sub cDestino_GotFocus()
    cDestino.SelStart = 0: cDestino.SelLength = Len(cDestino.Text)
End Sub

Private Sub cDestino_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If CargoDatos Then Foco tFecha
    End If
End Sub

Private Sub cOrigen_Click()
    vsLista.Rows = 1
    cDestino.Text = ""
End Sub

Private Sub cOrigen_GotFocus()
    cOrigen.SelStart = 0: cOrigen.SelLength = Len(cOrigen.Text)
End Sub

Private Sub cOrigen_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn And cOrigen.ListIndex <> -1 Then
        On Error GoTo errBusco
        Screen.MousePointer = 11
        
        cons = "Select * from Moneda Where MonCodigo = " & cOrigen.ItemData(cOrigen.ListIndex)
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            If Not IsNull(rsAux!MonCotizaEn) Then BuscoCodigoEnCombo cDestino, rsAux!MonCotizaEn
        End If
        rsAux.Close
        
        If cDestino.ListIndex <> -1 Then
            If CargoDatos Then Foco tFecha
        Else
            Foco cDestino
        End If
        
        Screen.MousePointer = 0
    End If
    Exit Sub

errBusco:
    clsGeneral.OcurrioError "Error al buscar la cotización para la moneda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub cOrigen_LostFocus()
    cOrigen.SelStart = 0
End Sub


Private Sub dBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        CargoDatos porBusqueda:=True
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub

Private Sub Form_Load()
On Error GoTo errLoad

    Me.BackColor = Colores.clVerde
'    Me.Caption = ""
    'ObtengoSeteoForm Me
    
    InicializoControles
    
    vsLista.Rows = 1
    TipoIngreso
    
    If paMonedaDolar <> 0 Then
        BuscoCodigoEnCombo cOrigen, CLng(paMonedaDolar)
        Call cOrigen_KeyPress(vbKeyReturn)
    End If
    
    CentroForm Me
    Exit Sub
    
errLoad:
    clsGeneral.OcurrioError "Error al iniciar el formulario.", Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    'GuardoSeteoForm Me
    CierroConexion
    Set clsGeneral = Nothing
End Sub

Private Sub Label1_Click()
    Foco cOrigen
End Sub

Private Sub Label2_Click()
    Foco cDestino
End Sub

Private Sub Label3_Click()
    Foco tFecha
End Sub

Private Sub Label4_Click()
    Foco tComprador
End Sub

Private Sub Label5_Click()
    Foco tVendedor
End Sub

Private Sub tComprador_GotFocus()
    tComprador.SelStart = 0: tComprador.SelLength = Len(tComprador.Text)
End Sub

Private Sub tComprador_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CanceloIngreso
End Sub

Private Sub tComprador_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(tComprador.Text) = "" Then Exit Sub
        If Not IsNumeric(tComprador.Text) Then Exit Sub
        Foco tVendedor
    End If
End Sub

Private Sub tComprador_LostFocus()
    If IsNumeric(tComprador.Text) Then tComprador.Text = Format(tComprador.Text, "#,##0.000") Else tComprador.Text = ""
    tComprador.SelStart = 0
End Sub

Private Sub tFecha_GotFocus()
    tFecha.SelStart = 0: tFecha.SelLength = Len(tFecha.Text)
End Sub

Private Sub tFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CanceloIngreso
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And IsDate(tFecha.Text) Then Foco tComprador
End Sub

Private Sub tFecha_LostFocus()
    tFecha.SelStart = 0
    If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy") Else tFecha.Text = ""
End Sub

Private Sub tVendedor_GotFocus()
    tVendedor.SelStart = 0: tVendedor.SelLength = Len(tVendedor.Text)
End Sub

Private Sub tVendedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then CanceloIngreso
End Sub

Private Sub tVendedor_KeyPress(KeyAscii As Integer)

    On Error Resume Next
    If KeyAscii = vbKeyReturn And IsNumeric(tVendedor.Text) Then
        tVendedor.Text = Format(tVendedor.Text, "#,##0.000")
        AccionGrabar
    End If
End Sub

Private Sub tVendedor_LostFocus()
    tVendedor.SelStart = 0
    If IsNumeric(tVendedor.Text) Then tVendedor.Text = Format(tVendedor.Text, "#,##0.000") Else tVendedor.Text = ""
End Sub


Private Sub TipoIngreso(Optional bModificar As Boolean = False)
Dim bkColor As Long

    If bModificar Then bkColor = Inactivo Else bkColor = vbWindowBackground
    
    cCotizacion.Enabled = Not bModificar
    cOrigen.Enabled = Not bModificar
    cDestino.Enabled = Not bModificar
    
    vsLista.Enabled = Not bModificar
    
    cCotizacion.BackColor = bkColor
    cOrigen.BackColor = bkColor
    cDestino.BackColor = bkColor
    
End Sub

Private Function ValidoCampos() As Boolean

    On Error GoTo errValido
    ValidoCampos = False
    
    If cCotizacion.ListIndex = -1 Then Foco cCotizacion: Exit Function
    If cOrigen.ListIndex = -1 Then Foco cOrigen: Exit Function
    If cDestino.ListIndex = -1 Then Foco cDestino: Exit Function
    
    If Not IsDate(tFecha.Text) Then Foco tFecha: Exit Function
    If Not IsNumeric(tComprador.Text) Then Foco tComprador: Exit Function
    If Not IsNumeric(tVendedor.Text) Then Foco tVendedor: Exit Function
    
    If CCur(tVendedor.Text) < CCur(tComprador.Text) Then
        MsgBox "El valor comprador debe ser menor al valor vendedor.", vbExclamation, "Posible Error"
        Foco tComprador: Exit Function
    End If
    
    Dim mTCOld As Currency
    Dim mVar As Currency
    mTCOld = modComun.TasadeCambio(cOrigen.ItemData(cOrigen.ListIndex), cDestino.ItemData(cDestino.ListIndex), CDate(tFecha.Text), _
                    TipoTC:=cCotizacion.ItemData(cCotizacion.ListIndex))
                 
    mVar = ((CCur(tComprador.Text) * 100) / mTCOld) - 100
    If mVar > 5 Then
        If MsgBox("La cotización de la moneda aumentó más de un 5% con respecto al último valor ingresado " & Format(mTCOld, "(#,##0.000)") & vbCrLf & _
                    "El valor ingresado, es el correcto ?", vbExclamation + vbYesNo + vbDefaultButton2, "La Cotización Aumentó más de un 5%") = vbNo Then
            Foco tComprador: Exit Function
        End If
    End If
    
    If mVar < -1 Then
        If MsgBox("La cotización de la moneda disminuyó más de un 1% con respecto al último valor ingresado " & Format(mTCOld, "(#,##0.000)") & vbCrLf & _
                    "El valor ingresado, es el correcto ?", vbExclamation + vbYesNo + vbDefaultButton2, "La Cotización Disminuyó más de un 1%") = vbNo Then
            Foco tComprador: Exit Function
        End If
    End If
    
    ValidoCampos = True
    Exit Function
    
errValido:
    clsGeneral.OcurrioError "Error al validar los datos.", Err.Description
End Function

Private Sub AccionGrabar()

    If Not ValidoCampos Then Exit Sub
    
    Dim aMsg As String
    
    aMsg = Format(tFecha.Text, "Long Date") & vbCrLf & _
                "Tipo de Cotización: " & Trim(cCotizacion.Text) & vbCrLf & vbCrLf & _
                Trim(cOrigen.Text) & " cotizados en " & Trim(cDestino.Text) & vbCrLf & _
                "  " & tComprador.Text & " Compra " & vbCrLf & _
                "  " & tVendedor.Text & " Venta " & vbCrLf & vbCrLf & _
                "Confirma grabar la cotización ingresada ?."
    
    If MsgBox(aMsg, vbQuestion + vbYesNo, "Grabar Cotización") = vbNo Then Exit Sub
    
    On Error GoTo ErrGrabar
    
    Dim mTipoC As Integer, mOrigen As Integer, MDestino As Integer
    Dim mFecha As String
    
    mTipoC = cCotizacion.ItemData(cCotizacion.ListIndex)
    mOrigen = cOrigen.ItemData(cOrigen.ListIndex)
    MDestino = cDestino.ItemData(cDestino.ListIndex)
    mFecha = Trim(tFecha.Text)
    
    cons = "Select * From TasaCambio " & _
               " Where TCaFecha = '" & Format(mFecha, "mm/dd/yyyy") & "'" & _
               " And TCaTipo = " & mTipoC & _
               " And TCaOriginal = " & mOrigen & _
               " And TCaDestino = " & MDestino
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsAux.EOF Then
        aMsg = "Existe una cotización '" & Trim(cCotizacion.Text) & "' en " & Trim(cOrigen.Text) & ", para la fecha ingresada." & vbCrLf & vbCrLf & _
                "La cotización actual es de " & vbCrLf & _
                "  " & Format(rsAux!TCaComprador, "#,##0.000") & " Compra " & vbCrLf & _
                "  " & Format(rsAux!TCaVendedor, "#,##0.000") & " Venta " & vbCrLf & vbCrLf & _
                "Ud. quiere actualizar la información ?."
                
        If MsgBox(aMsg, vbQuestion + vbYesNo + vbDefaultButton2, "Cotización Ingresada !!!") = vbNo Then
            rsAux.Close
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    
    
    If rsAux.EOF Then rsAux.AddNew Else rsAux.Edit
    
    rsAux!TCaTipo = mTipoC
    rsAux!TCaFecha = Format(mFecha, "mm/dd/yyyy")
    rsAux!TCaOriginal = mOrigen
    rsAux!TCaDestino = MDestino
    
    rsAux!TCaComprador = CCur(tComprador.Text)
    rsAux!TCaVendedor = CCur(tVendedor.Text)
                   
    rsAux.Update
    rsAux.Close
    
  
    CargoDatos
    
    On Error Resume Next
    If Weekday(tFecha.Text) = 6 Then tFecha.Text = CDate(tFecha.Text) + 3 Else tFecha.Text = CDate(tFecha.Text) + 1
    tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy")
    tComprador.Text = ""
    tVendedor.Text = ""
    Foco tFecha
        
    Exit Sub
    
ErrGrabar:
    clsGeneral.OcurrioError "Error al grabar los datos de la cotización.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function AccionEliminar() As Boolean
    On Error GoTo errEliminar
    
    If vsLista.Rows = 1 Then Exit Function
    If cCotizacion.ListIndex = -1 Then Foco cCotizacion: Exit Function
    If cOrigen.ListIndex = -1 Then Foco cOrigen: Exit Function
    If cDestino.ListIndex = -1 Then Foco cDestino: Exit Function
    
    Dim mFecha As String
    mFecha = vsLista.Cell(flexcpText, vsLista.Row, 0)
    
    Dim mCotizacion As Integer
    Dim mOrigen As Integer, MDestino As Integer
    
    mCotizacion = cCotizacion.ItemData(cCotizacion.ListIndex)
    mOrigen = cOrigen.ItemData(cOrigen.ListIndex)
    MDestino = cDestino.ItemData(cDestino.ListIndex)
    
    Dim aMsg As String
    
    aMsg = Format(mFecha, "Long Date") & vbCrLf & _
                "Tipo de Cotización: " & Trim(cCotizacion.Text) & vbCrLf & vbCrLf & _
                Trim(cOrigen.Text) & " cotizados en " & Trim(cDestino.Text) & vbCrLf & _
                "  " & vsLista.Cell(flexcpText, vsLista.Row, 1) & " Compra " & vbCrLf & _
                "  " & vsLista.Cell(flexcpText, vsLista.Row, 2) & " Venta " & vbCrLf & vbCrLf & _
                "Confirma eliminar la cotización seleccionada ?."
                
    If MsgBox(aMsg, vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar Cotización") = vbNo Then Exit Function
    
    Screen.MousePointer = 11
    cons = "Select * From TasaCambio " & _
               " Where TCaFecha = '" & Format(mFecha, "mm/dd/yyyy") & "'" & _
               " And TCaTipo = " & mCotizacion & _
               " And TCaOriginal = " & mOrigen & _
               " And TCaDestino = " & MDestino
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then rsAux.Delete
    rsAux.Close
        
    vsLista.RemoveItem vsLista.Row
    Screen.MousePointer = 0
    Exit Function

errEliminar:
    clsGeneral.OcurrioError "Error al eliminar la cotización.", Err.Description
    Screen.MousePointer = 0
End Function

Private Function CargoDatos(Optional porBusqueda As Boolean = False) As Boolean
On Error GoTo ErrCD
    
    CargoDatos = False
    vsLista.Rows = 1
    
    If cCotizacion.ListIndex = -1 Then Foco cCotizacion: Exit Function
    If cOrigen.ListIndex = -1 Then Foco cOrigen: Exit Function
    If cDestino.ListIndex = -1 Then Foco cDestino: Exit Function
    
    CargoDatos = True
    Dim mCotizacion As Integer
    Dim mOrigen As Integer, MDestino As Integer
    
    mCotizacion = cCotizacion.ItemData(cCotizacion.ListIndex)
    mOrigen = cOrigen.ItemData(cOrigen.ListIndex)
    MDestino = cDestino.ItemData(cDestino.ListIndex)
    
    Screen.MousePointer = 11
    
    If Not porBusqueda Then
        cons = "Select Top 30 * From TasaCambio "
    Else
        cons = "Select * From TasaCambio "
    End If
    
    cons = cons & " Where TCaTipo = " & mCotizacion & _
                           " And TCaOriginal = " & mOrigen & _
                           " And TCaDestino = " & MDestino
                           
    If porBusqueda Then
        Dim mFecha1 As String, mFecha2 As String
        mFecha1 = dBuscar.Value - 10
        mFecha2 = dBuscar.Value + 10
        
        cons = cons & " And TCaFecha Between " & Format(mFecha1, "'mm/dd/yyyy'") & _
                                                           " And " & Format(mFecha2, "'mm/dd/yyyy'")
    End If
    
    cons = cons & " Order by TCaFecha Desc"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not rsAux.EOF
        With vsLista
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Format(rsAux!TCaFecha, "dd/mm/yyyy")
            .Cell(flexcpText, .Rows - 1, 1) = Format(rsAux!TCaComprador, "#,##0.000")
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!TCaVendedor, "#,##0.000")
            
        End With
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    On Error Resume Next
    If porBusqueda Then
        Dim I As Long, mValor As String
        mValor = dBuscar.Value
        
        For I = 1 To vsLista.Rows - 1
            If mValor = vsLista.Cell(flexcpText, I, 0) Then
                vsLista.Select I, 0, I, 0
                Exit For
            End If
        Next
    Else
        vsLista.Select 1, 0, 0, 0
    End If
    
    If vsLista.Rows > 1 Then vsLista.SetFocus
    Screen.MousePointer = 0
    Exit Function
    
ErrCD:
    clsGeneral.OcurrioError "Error al cargar la lista.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub InicializoControles()
    
    On Error Resume Next
    
    cons = "Select MonCodigo, MonNombre From Moneda Order by MonSigno"
    CargoCombo cons, cOrigen
    CargoCombo cons, cDestino
    
    cons = "Select TCoCodigo, TCoNombre from TipoCotizacion Order by TCoNombre"
    CargoCombo cons, cCotizacion
    BuscoCodigoEnCombo cCotizacion, 1
    
    With vsLista
        .Editable = False
        .Rows = 1: .Cols = 3
        .FormatString = "<Fecha|>Comprador|>Vendedor"
        .ColWidth(0) = 1000
        .ColWidth(1) = 950
        .ColWidth(2) = 900
        .AllowUserResizing = flexResizeColumns
        '.GridLines = flexGridFlatHorz
        .ExtendLastCol = True
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorBkg = Me.BackColor
        .BackColor = Colores.Blanco
        .BackColorAlternate = &HC0FFC0   'Colores.clNaranja
        .BorderStyle = flexBorderNone
        .HighLight = flexHighlightWithFocus
        .RowHeight(0) = 300
    End With
        
   bExit.BackColor = Me.BackColor
   vsLista.BackColorFixed = Me.BackColor
   
End Sub

Private Sub vsLista_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    If vsLista.Rows = 1 Then Exit Sub
    
    Select Case KeyCode
        Case vbKeySpace
            With vsLista
                tFecha.Text = .Cell(flexcpText, .Row, 0)
                tComprador.Text = .Cell(flexcpText, .Row, 1)
                tVendedor.Text = .Cell(flexcpText, .Row, 2)
            End With
            Foco tComprador
        
        Case vbKeyDelete: AccionEliminar
    End Select
    
End Sub

Private Function CanceloIngreso()
    tFecha.Text = ""
    tComprador.Text = ""
    tVendedor.Text = ""
    Foco tFecha
End Function
