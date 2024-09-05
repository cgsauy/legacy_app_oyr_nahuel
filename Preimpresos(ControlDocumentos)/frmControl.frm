VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{191D08B9-4E92-4372-BF17-417911F14390}#1.5#0"; "orGridPreview.ocx"
Begin VB.Form frmControl 
   Appearance      =   0  'Flat
   Caption         =   "Control de Documentos"
   ClientHeight    =   6735
   ClientLeft      =   2190
   ClientTop       =   1545
   ClientWidth     =   8625
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmControl.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   8625
   Begin VB.CommandButton butHelp 
      Caption         =   "Ayuda"
      Height          =   315
      Left            =   5040
      TabIndex        =   21
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton butBuscarNroRojo 
      Caption         =   "Buscar #Rojo"
      Height          =   315
      Left            =   6360
      TabIndex        =   11
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtSerieRojo 
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   5
      ToolTipText     =   "Serie"
      Top             =   540
      Width           =   255
   End
   Begin VB.CommandButton bAccion 
      Caption         =   "Editar"
      Height          =   320
      Index           =   5
      Left            =   6360
      TabIndex        =   20
      Top             =   4200
      Width           =   1215
   End
   Begin VB.ComboBox cTipo 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   180
      Width           =   915
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "Imprimir"
      Height          =   315
      Left            =   6360
      TabIndex        =   19
      Top             =   5340
      Width           =   1215
   End
   Begin VB.CommandButton bAccion 
      Caption         =   "Anu && Ext"
      Height          =   320
      Index           =   4
      Left            =   6360
      TabIndex        =   18
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton bAccion 
      Caption         =   "Extraviado"
      Height          =   320
      Index           =   3
      Left            =   6360
      TabIndex        =   17
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton bGrabar 
      Caption         =   "Grabar"
      Height          =   315
      Left            =   6360
      TabIndex        =   16
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton bAccion 
      Caption         =   "Reubicar"
      Height          =   320
      Index           =   2
      Left            =   6360
      TabIndex        =   15
      Top             =   3540
      Width           =   1215
   End
   Begin VB.CommandButton bAccion 
      Caption         =   "Limpiar Edo."
      Height          =   320
      Index           =   1
      Left            =   6360
      TabIndex        =   14
      Top             =   2700
      Width           =   1215
   End
   Begin VB.CommandButton bAccion 
      Caption         =   "Anulado"
      Height          =   320
      Index           =   0
      Left            =   6360
      TabIndex        =   13
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox tRojoH 
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   5340
      MaxLength       =   10
      TabIndex        =   8
      Top             =   540
      Width           =   915
   End
   Begin VB.TextBox tRojoD 
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   6
      Top             =   540
      Width           =   855
   End
   Begin VB.ComboBox cPreimpreso 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   3255
   End
   Begin VB.CommandButton bCancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   6360
      TabIndex        =   12
      Top             =   6180
      Width           =   1215
   End
   Begin VB.CommandButton bCargar 
      Caption         =   "&Cargar Lista"
      Height          =   315
      Left            =   6360
      TabIndex        =   9
      Top             =   540
      Width           =   1215
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
      Height          =   5055
      Left            =   60
      TabIndex        =   10
      Top             =   960
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   8916
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483645
      ForeColorFixed  =   -2147483634
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14737632
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   8421631
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   4
      GridLinesFixed  =   5
      GridLineWidth   =   1
      Rows            =   2
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
   Begin MSComCtl2.DTPicker tFecha 
      Height          =   315
      Left            =   960
      TabIndex        =   3
      Top             =   540
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   46399489
      CurrentDate     =   37543
   End
   Begin orGridPreview.GridPreview cPrint 
      Left            =   7260
      Top             =   0
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lHasta 
      BackStyle       =   0  'Transparent
      Caption         =   "&hasta:"
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lArranca 
      BackStyle       =   0  'Transparent
      Caption         =   "&Arranca en:"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   600
      Width           =   915
   End
   Begin VB.Label lFecha 
      BackStyle       =   0  'Transparent
      Caption         =   "&Fecha:"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   615
      Width           =   615
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "?"
      Visible         =   0   'False
      Begin VB.Menu MnuHlp 
         Caption         =   "Ayuda"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum cols
    Rojo = 0
    TipoDoc
    Numero
    Estado
End Enum

Private Type typData
    txtTipoDoc As String
    txtDocumento As String
    idTipoDoc As Integer
    idEstado As Byte
End Type

Dim dFila1 As typData
Dim dFila2 As typData

Private prmIDPreimpreso As Integer
Private prmTipoDocs As String
Private prmSucursal  As Integer
Private prmFecha As Date

Private prmROJO_ASC As Boolean

Private Sub bCancel_Click()
    Unload Me
End Sub

Private Sub bCargar_Click()
    fnc_CargoListas
End Sub

Private Sub bGrabar_Click()
    AccionGrabar
End Sub

Private Sub bAccion_Click(Index As Integer)
On Error GoTo errACC
    If Not (vsLista.Rows > vsLista.FixedRows) Then Exit Sub
    
    Select Case Index
        Case 0: fnc_Anulado xEstado:=1     'Papel Anulado
        Case 3: fnc_Anulado xEstado:=10    'Extraviado
        Case 4: fnc_Anulado xEstado:=11    'Extraviado & Anulado
        
        Case 5: fnc_EditarDatos
        
        Case 1: fnc_NoEsAnulado
                
        Case 2
                If Val(bAccion(Index).Tag) = 0 Then
                    bAccion(Index).Tag = vsLista.Row
                    vsLista.Cell(flexcpForeColor, vsLista.Row, 0, , vsLista.cols - 1) = vsLista.ForeColorSel
                    vsLista.Cell(flexcpBackColor, vsLista.Row, 0, , vsLista.cols - 1) = vsLista.BackColorSel
                    zfn_StateReubicar xSI:=True
                Else
                    vsLista.Cell(flexcpForeColor, Val(bAccion(Index).Tag), 0, , vsLista.cols - 1) = vsLista.ForeColor
                    vsLista.Cell(flexcpBackColor, Val(bAccion(Index).Tag), 0, , vsLista.cols - 1) = vsLista.BackColor

                    fnc_CambiarF1xF2 Val(bAccion(Index).Tag), vsLista.Row

                    bAccion(Index).Tag = 0
                    zfn_StateReubicar xSI:=False
                End If
    End Select
    
    Exit Sub
    
errACC:
    clsGeneral.OcurrioError "Error al procesar la acción solicitada.", Err.Description
End Sub


Private Function zfn_StateReubicar(xSI As Boolean)
    cTipo.Enabled = Not xSI: cPreimpreso.Enabled = Not xSI
    lFecha.Enabled = Not xSI: tFecha.Enabled = Not xSI
    lArranca.Enabled = Not xSI: tRojoD.Enabled = Not xSI
    lHasta.Enabled = Not xSI: tRojoH.Enabled = Not xSI
    txtSerieRojo.Enabled = tRojoD.Enabled
    
    bCargar.Enabled = Not xSI
    bAccion(0).Enabled = Not xSI
    bAccion(1).Enabled = Not xSI
    bAccion(4).Enabled = Not xSI
    bAccion(3).Enabled = Not xSI
    bGrabar.Enabled = Not xSI
    bCancel.Enabled = Not xSI
    
End Function

Private Sub bPrint_Click()
    
    With cPrint
        .Orientation = opPortrait
        .Caption = Me.Caption
        .Header = Me.Caption & " - " & cPreimpreso.Text & IIf(cTipo.ListIndex = 0, " al " & tFecha.Value, "")
        .PageBorder = opTopBottom
        .MarginLeft = 1200
        .MarginTop = 800
    End With
    
    Screen.MousePointer = 11
    vsLista.ExtendLastCol = False
    
    With cPrint
        .Columns = 3
        .AddGrid vsLista.hwnd
        .LineAfterGrid ""
        .LineAfterGrid "Glosario Estados: PA=Papel Anulado; EXT=Extraviado; ExA=Extraviado y anulado"
        .ShowPreview
    End With
    
    vsLista.ExtendLastCol = True
    Screen.MousePointer = 0

End Sub


Private Sub butBuscarNroRojo_Click()
Dim adTexto As String
    adTexto = InputBox("Ingrese serie y número del documento", "Buscar número rojo")
    If Trim(adTexto) <> "" Then
        Dim mDSerie As String, mDNumero As Long
        If InStr(adTexto, "-") <> 0 Then
            mDSerie = Mid(adTexto, 1, InStr(adTexto, "-") - 1)
            mDNumero = Val(Mid(adTexto, InStr(adTexto, "-") + 1))
        Else
            adTexto = Replace(adTexto, " ", "")
            If IsNumeric(Mid(adTexto, 2, 1)) Then
                mDSerie = Mid(adTexto, 1, 1)
                mDNumero = Val(Mid(adTexto, 2))
            Else
                mDSerie = Mid(adTexto, 1, 2)
                mDNumero = Val(Mid(adTexto, 3))
            End If
        End If
        Cons = "SELECT CDoSerieRoja 'Serie Roja', CDoNumeroRojo 'Nro. Rojo', dbo.NombreTipoDocumento(CDoTipoDoc) Documento, CDoSerie Serie, CDoNumero Número " & _
                "FROM ControlDocumentos " & _
                "WHERE CDoSerie = '" & mDSerie & "' AND CDoNumero = " & mDNumero
        Dim objLA As New clsListadeAyuda
        objLA.ActivarAyuda cBase, Cons, 4400, 0, "Serie y Número rojo"
        Set objLA = Nothing
    End If
End Sub

Private Sub butHelp_Click()
    On Error GoTo errHelp
    Screen.MousePointer = 11
    
    Dim aFile As String
    Cons = "Select * from Aplicacion Where AplNombre = 'Control'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux!AplHelp) Then aFile = Trim(RsAux!AplHelp)
    RsAux.Close
    
    If aFile <> "" Then EjecutarApp aFile
    
    Screen.MousePointer = 0
    Exit Sub
    
errHelp:
    clsGeneral.OcurrioError "Error al activar el archivo de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub cPreimpreso_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If cPreimpreso.ListIndex <> -1 Then
            On Error GoTo errPP
            tRojoD.Text = "": tRojoH.Text = "": txtSerieRojo.Text = ""
            
            'mSQL = "Select Max(CDoNumeroRojo) from ControlDocumentos Where CDoPreimpreso = " & cPreimpreso.ItemData(cPreimpreso.ListIndex)
            mSQL = "Select TOP 1 IsNull(CDOSerieRoja, '0'), CDoNumeroRojo FROM ControlDocumentos Where CDoPreimpreso = " & cPreimpreso.ItemData(cPreimpreso.ListIndex) _
                & " ORDER BY CDoSerieRoja DESC, CDoNumeroRojo DESC"
            
            Set RsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                If Not IsNull(RsAux(1)) Then tRojoD.Text = RsAux(1) + 1
                If Not IsNull(RsAux(0)) Then txtSerieRojo.Text = Trim(RsAux(0))
            End If
            RsAux.Close
        
            If tFecha.Enabled Then tFecha.SetFocus Else txtSerieRojo.SetFocus
        End If
    End If
    
    Exit Sub
errPP:
    clsGeneral.OcurrioError "Error al buscar el último número de Preimpreso.", Err.Description
End Sub

Private Sub cTipo_Click()

Dim xSI As Boolean

    Select Case cTipo.ListIndex
        Case 0: xSI = False
        Case 1: xSI = True
    End Select

    lFecha.Enabled = Not xSI: tFecha.Enabled = Not xSI
    
    bAccion(0).Enabled = Not xSI
    bAccion(1).Enabled = Not xSI
    bAccion(2).Enabled = Not xSI
    bAccion(3).Enabled = Not xSI
    bAccion(4).Enabled = Not xSI
    bGrabar.Enabled = Not xSI
    
    bAccion(5).Enabled = True 'xSI
    vsLista.Rows = vsLista.FixedRows
    
End Sub

Private Sub cTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cPreimpreso.SetFocus
End Sub

Private Sub Form_Load()
    InicializoForm
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    On Error GoTo errRZ
    vsLista.Height = Me.ScaleHeight - vsLista.Top - 100
    
errRZ:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EndMain
End Sub

Private Sub InicializoForm()
    
    On Error Resume Next
    LimpioFicha
    
    cTipo.AddItem "Control", 0
    cTipo.AddItem "Listar", 1
    
    tFecha.Value = Date
    cTipo.ListIndex = 0
    
    mSQL = "Select PreCodigo, PreNombre from Preimpreso Order by PreNombre  "
    CargoCombo mSQL, cPreimpreso
    
    With vsLista
        .Rows = 1: .cols = 1
        .FormatString = "<Nº Rojo|<Docum|<Número|<Estado|"
        .ColWidth(cols.Rojo) = 900
        .ColWidth(cols.Estado) = 500
        .WordWrap = False
        .MergeCells = flexMergeSpill
        .ExtendLastCol = True
        
        .Editable = False
        .RowHeight(0) = 280
        .SelectionMode = flexSelectionByRow
    End With

End Sub

Private Function fnc_UltimoGrabados(tb_Documento As Long, tb_Traslado As Long, tb_Resguardos As Date)
Dim mIDDocumento As Long, mIDTRaslado As Long

    mIDDocumento = 0: mIDTRaslado = 0
    
    'Busco el Minimo Documento para esa fecha ------------------------------------
    mSQL = "Select Min(DocCodigo) From Documento " & _
                " Where DocTipo IN (" & prmTipoDocs & ")" & _
                " And DocSucursal = " & prmSucursal & _
                " And DocFecha Between '" & Format(prmFecha, "mm/dd/yyyy 00:00:00") & "' And '" & Format(prmFecha, "mm/dd/yyyy 23:59:59") & "'" & _
                " AND DocCodigo NOT IN(SELECT TAIDocumento FROM TicketsAImprimir)"
                
    'InputBox "SQL Menor Código para la fecha:", "Confirmar 1 ", mSQL

    Set RsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux(0)) Then mIDDocumento = RsAux(0)
    End If
    RsAux.Close
    
    Dim arrTMP() As String, iX As Integer, mSTR As String
    arrTMP = Split(prmTipoDocs, ",")
    For iX = LBound(arrTMP) To UBound(arrTMP)
        mSTR = mSTR & IIf(mSTR = "", "", " OR ") & "DocTipo = " & arrTMP(iX)
    Next
    
    If mIDDocumento <> 0 Then
        'DocTipo IN (" & prmTipoDocs & ")"
        mSQL = "Select TOP 1 * from ControlDocumentos, Documento" & _
                    " Where CDoPreimpreso = " & prmIDPreimpreso & _
                    " And CDoTipoDoc = DocTipo And CDoSerie = DocSerie And CDoNumero = DOcNumero " & _
                    " AND (" & mSTR & ")  And DocSucursal = " & prmSucursal & _
                    " And DocFecha Between '" & Format(prmFecha, "mm/dd/yyyy 00:00:00") & "' And '" & Format(prmFecha, "mm/dd/yyyy 23:59:59") & "'" & _
                    " Order by DocFecha DESC, DocCodigo DESC"
        
        'InputBox "SQL Saco el id del documento para los tipos y fechas:", "Confirmar 2 ", mSQL
        
        Set RsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then mIDDocumento = RsAux!DocCodigo Else mIDDocumento = 0
        RsAux.Close
    End If

    tb_Documento = mIDDocumento
    
    If InStr(prmTipoDocs, TipoDocumento.Traslados) <> 0 Then    'Tabla TRASLADO -----------------------------------------------
                
        mSQL = "Select Min(TraNumero) From Traspaso " & _
                    " Where TraSucursal = " & prmSucursal & _
                    " And TraFImpreso Between '" & Format(prmFecha, "mm/dd/yyyy 00:00:00") & "' And '" & Format(prmFecha, "mm/dd/yyyy 23:59:59") & "'"
        
        'InputBox "SQL Menor ID de Traslados :", "Confirmar 3 ", mSQL
        
        Set RsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            If Not IsNull(RsAux(0)) Then mIDTRaslado = RsAux(0)
        End If
        RsAux.Close
        
        If mIDTRaslado <> 0 Then
            mSQL = "Select TOP 1 * from ControlDocumentos, Traspaso" & _
                        " Where CDoPreimpreso = " & prmIDPreimpreso & _
                        " And CDoSerie = TraSerie And CDoNumero = TraNumero " & _
                        " AND CDoTipoDoc = " & TipoDocumento.Traslados & " And TraSucursal = " & prmSucursal & _
                        " And TraFImpreso Between '" & Format(prmFecha, "mm/dd/yyyy 00:00:00") & "' And '" & Format(prmFecha, "mm/dd/yyyy 23:59:59") & "'" & _
                        " Order by TraFImpreso DESC, TraCodigo DESC"
            
            'InputBox "SQL Traslados:", "Confirmar 4 ", mSQL
            
            Set RsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then mIDTRaslado = RsAux!Tranumero Else mIDTRaslado = 0
            RsAux.Close
        End If

    End If      '---------------------------------------------------------------------------------------------------------------------------------
    tb_Traslado = mIDTRaslado
    
    'Resguardo -----------------------------------------------------------------
    Dim mFechaResguardo As Date
    If prmSucursal = 5 Then
        mSTR = ""
        For iX = LBound(arrTMP) To UBound(arrTMP)
            mSTR = mSTR & IIf(mSTR = "", "", " OR ") & "TDoID = " & arrTMP(iX)
        Next
        
        'Dim mIDResguardo As Long
        mFechaResguardo = DateSerial(1990, 1, 1)
        mSQL = "Select Min(ComFechaModificacion) From ZureoCGSA.dbo.cceComprobantes, TipoDocumento " & _
                    " WHERE ComFechaModificacion Between '" & Format(prmFecha, "mm/dd/yyyy 00:00:00") & "' And '" & Format(prmFecha, "mm/dd/yyyy 23:59:59") & "'" & _
                    " AND ComTipo = TDoTipoDocZureo AND (" & mSTR & ")"
        
        'InputBox "SQL Comprobantes Zureo :", "Confirmar 4 ", mSQL
        
        Set RsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            If Not IsNull(RsAux(0)) Then mFechaResguardo = RsAux(0)
        End If
        RsAux.Close
            
        If mFechaResguardo > DateSerial(1990, 1, 1) Then
        
            mFechaResguardo = prmFecha
        
            mSQL = "Select  Max(ComFechaModificacion) from ControlDocumentos, ZureoCGSA.dbo.cceComprobantes, TipoDocumento " & _
                        "WHERE CDoPreimpreso = " & prmIDPreimpreso & _
                        " And CDoSerie = CASE CharIndex('-',ComNumero) WHEN 0 THEN '?' ELSE SubString(ComNumero, 1, CharIndex('-',ComNumero)-1) COLLATE Modern_Spanish_CI_AI END " & _
                        " And CDoNumero = CASE CharIndex('-',ComNumero) WHEN 0 THEN ComID ELSE CONVERT(int, SubString(ComNumero, CharIndex('-',ComNumero)+1,6)) END " & _
                        " AND CDoTipoDoc = TDoID AND ComTipo = TDoTipoDocZureo" & _
                        " AND (" & mSTR & ")" & _
                        " AND ComFechaModificacion Between '" & Format(prmFecha, "mm/dd/yyyy 00:00:00") & "' And '" & Format(prmFecha, "mm/dd/yyyy 23:59:59") & "'"
                        '& _
                        " Order by ComFechaModificacion DESC, ComID DESC"
            
            'InputBox "SQL Comprobantes Zureo :", "Confirmar 5 ", mSQL
            
            Set RsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then If Not IsNull(RsAux(0)) Then mFechaResguardo = RsAux(0)
            RsAux.Close
        End If
    End If
    tb_Resguardos = mFechaResguardo
    'Resguardo -----------------------------------------------------------------
    
    
End Function

Private Function fnc_CargoListas()
On Error GoTo errCL

    If cPreimpreso.ListIndex = -1 Then cPreimpreso.SetFocus: Exit Function
    If Trim(txtSerieRojo.Text) = "" Then txtSerieRojo.SetFocus: Exit Function
    If Not IsNumeric(tRojoD.Text) And Not IsNumeric(tRojoH.Text) Then
        tRojoD.SetFocus: Exit Function
    Else
        If CLng(tRojoD.Text) > CLng(tRojoH.Text) Then
            MsgBox "Desde es mayor a hasta.", vbExclamation, "Atención"
            Exit Function
        End If
    End If

    prmFecha = tFecha.Value
    prmIDPreimpreso = cPreimpreso.ItemData(cPreimpreso.ListIndex)
    
    mSQL = "Select * from Preimpreso Where PreCodigo = " & prmIDPreimpreso
    Set RsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        prmTipoDocs = Trim(RsAux!PreTipoDocs)
        prmSucursal = Trim(RsAux!PreSucursal)
    End If
    RsAux.Close

    If prmTipoDocs = "" Or prmSucursal = 0 Then
        MsgBox "No hay datos para cargar los documentos emitidos, faltan los tipos asociados al preimpreso o el código de scursal.", vbExclamation, "Faltan datos"
        Exit Function
    End If
    '----------------------    ----------------------   ----------------------  ----------------------  ----------------------
    
    Screen.MousePointer = 11
    vsLista.Rows = vsLista.FixedRows
    
    If cTipo.ListIndex = 0 Then fnc_CargoParaControl
    If cTipo.ListIndex = 1 Then fnc_CargoGrabados
    
    bAccion(5).Enabled = (cTipo.ListIndex = 1)
    
    Screen.MousePointer = 0
    Exit Function
errCL:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar la lista de documentos.", Err.Description
End Function

Private Function fnc_CargoParaControl()

    Dim xROJO As Long, xValor As Long
        
    Dim xQTotal As Long: xQTotal = -1
    
6    If IsNumeric(tRojoD.Text) And IsNumeric(tRojoH.Text) Then
        prmROJO_ASC = (Val(tRojoH.Text) >= Val(tRojoD.Text))
        xQTotal = IIf(prmROJO_ASC, Val(tRojoH.Text) - Val(tRojoD.Text), Val(tRojoD.Text) - Val(tRojoH.Text)) + 1
        xROJO = Val(tRojoD.Text)
    Else
        If IsNumeric(tRojoD.Text) Then
            xROJO = Val(tRojoD.Text): prmROJO_ASC = True
        Else
            xROJO = Val(tRojoH.Text): prmROJO_ASC = False
        End If
    End If
    
    
    Dim desde_TableDoc As Long, desde_TableTraspaso As Long, desde_TableResguardo As Date
    
    fnc_UltimoGrabados desde_TableDoc, desde_TableTraspaso, desde_TableResguardo
    
    'mSQL = "Select " & IIf(xQTotal <> -1, " TOP " & xQTotal, "") & " DocTipo, DocSerie, DocNumero From Documento " & _
                " Where DocTipo IN (" & prmTipoDocs & ")" & _
                " And DocSucursal = " & prmSucursal & _
                " And DocFecha Between '" & Format(prmFecha, "mm/dd/yyyy 00:00:00") & "' And '" & Format(prmFecha, "mm/dd/yyyy 23:59:59") & "'" & _
                " And DocCodigo > " & desde_TableDoc & _
                " Order by DocFecha, DocCodigo"
                
    mSQL = "Select  DocTipo, DocSerie, DocNumero, DocFecha, DocCodigo From Documento (index = iTipoFechaSucursalMoneda) " & _
                " Where DocTipo IN (" & prmTipoDocs & ")" & _
                " And DocSucursal = " & prmSucursal & _
                " And DocFecha Between '" & Format(prmFecha, "mm/dd/yyyy 00:00:00") & "' And '" & Format(prmFecha, "mm/dd/yyyy 23:59:59") & "'" & _
                " And NOT EXISTS (SELECT * FROM ControlDocumentos (INDEX = iPreimpTipodocSerNum) WHERE CDoPreimpreso = " & cPreimpreso.ItemData(cPreimpreso.ListIndex) & _
                " AND DocTipo = CDoTipoDoc AND DocSerie = CDoSerie AND DocNumero = CDoNumero)" & _
                " AND DocCodigo NOT IN(SELECT TAIDocumento FROM TicketsAImprimir)"
                
'IIf(desde_TableDoc <> 0, " And DocCodigo > " & desde_TableDoc, "") & _

    If InStr(prmTipoDocs, TipoDocumento.Traslados) <> 0 Then
        mSQL = mSQL & " UNION ALL " & _
                " Select  20 as DocTipo, TraSerie as DocSerie, TraNumero as DocNumero, TraFImpreso as DocFecha, TraCodigo as DocCodigo From Traspaso (index = iNumeroDoc)" & _
                " Where TraSucursal = " & prmSucursal & _
                " And TraFImpreso Between '" & Format(prmFecha, "mm/dd/yyyy 00:00:00") & "' And '" & Format(prmFecha, "mm/dd/yyyy 23:59:59") & "'" & _
                IIf(desde_TableTraspaso <> 0, " And TraNumero > " & desde_TableTraspaso, "")
    End If

    'RESGUARDOS
    mSQL = mSQL & " UNION ALL " & _
        "SELECT TDoID DocTipo, CASE CharIndex('-',ComNumero) WHEN 0 THEN '?' ELSE SubString(ComNumero, 1, CharIndex('-',ComNumero)-1) COLLATE Modern_Spanish_CI_AI END DocSerie, " & _
        "CASE CharIndex('-',ComNumero) WHEN 0 THEN ComID ELSE CONVERT(int, SubString(ComNumero, CharIndex('-',ComNumero)+1,6)) END DocNumero, ComFechaModificacion DocFecha, ComID DocCodigo " & _
        "FROM ZureoCGSA.dbo.cceComprobantes INNER JOIN TipoDocumento ON ComTipo = TDoTipoDocZureo " & _
        "WHERE ComFechaModificacion Between '" & Format(prmFecha, "mm/dd/yyyy 00:00:00") & "' And '" & Format(prmFecha, "mm/dd/yyyy 23:59:59") & "'" & _
        "AND TDoID IN (" & prmTipoDocs & ") AND ComFechaModificacion > '" & Format(desde_TableResguardo, "mm/dd/yyyy hh:nn:ss") & "' And IsNull(ComEstado, 0) <> 9 " & _
        "AND NOT EXISTS (SELECT * FROM ControlDocumentos WHERE CDoTipoDoc IN (41, 42) " & _
        "AND CASE CharIndex('-',ComNumero) WHEN 0 THEN '?' ELSE SubString(ComNumero, 1, CharIndex('-',ComNumero)-1) COLLATE Modern_Spanish_CI_AI END = CDoSerie " & _
        "AND CASE CharIndex('-',ComNumero) WHEN 0 THEN ComID ELSE CONVERT(int, SubString(ComNumero, CharIndex('-',ComNumero)+1,6)) END = CDoNumero)"

    mSQL = mSQL & " Order by DocFecha, DocCodigo"
    
    'Resguardos validar que el campo ComResguardo sea 1 o 2.
                
    'InputBox "SQL para cargar grilla:", "Confirmar", mSQL
    cBase.QueryTimeout = 90
    Set RsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    cBase.QueryTimeout = 20
    Dim xAdd As Long: xAdd = 0
    
    bGrabar.Enabled = (Not RsAux.EOF)
    
    Do While Not RsAux.EOF
        With vsLista
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, cols.Rojo) = xROJO
            .Cell(flexcpText, .Rows - 1, cols.TipoDoc) = RetornoNombreDocumento(RsAux!DocTipo, True)
            xValor = RsAux!DocTipo: .Cell(flexcpData, .Rows - 1, cols.TipoDoc) = xValor
            .Cell(flexcpData, .Rows - 1, 2) = txtSerieRojo.Text
            
            .Cell(flexcpText, .Rows - 1, cols.Numero) = Trim(RsAux!DocSerie) & "-" & RsAux!DocNumero
            xROJO = IIf(prmROJO_ASC, xROJO + 1, xROJO - 1)
            
            If xQTotal <> -1 Then
                xAdd = xAdd + 1
                If xAdd = xQTotal Then Exit Do
            End If
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close

End Function

Private Sub LimpioFicha()
    bGrabar.Enabled = False
End Sub

Private Function ValidoGrabar() As Boolean
On Error GoTo errValidar

    ValidoGrabar = False
        
   
    ValidoGrabar = True
    
errValidar:
End Function

Private Sub AccionGrabar()
On Error GoTo errGrabar
   
    If Not ValidoGrabar Then Exit Sub

    If MsgBox("¿Confirma grabar el control de los documentos?", vbQuestion + vbYesNo, "Control de Documentos") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    
    Dim mSQL As String, rsAdd As rdoResultset
    Dim arrDAT() As String
    Dim idX As Long
    
    mSQL = "Select * from ControlDocumentos Where CDoPreimpreso = 0 And CDoNumeroRojo = 0"
    Set rsAdd = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    With vsLista
        For idX = .FixedRows To .Rows - 1
            arrDAT = Split(.Cell(flexcpText, idX, cols.Numero), "-")
            rsAdd.AddNew
            
            rsAdd!CDoPreimpreso = prmIDPreimpreso
            rsAdd!CDONumeroRojo = .Cell(flexcpText, idX, cols.Rojo)
            rsAdd!CDoSerieRoja = txtSerieRojo.Text
            
            If Val(.Cell(flexcpData, idX, cols.TipoDoc)) <> 0 Then rsAdd!CDoTipoDoc = Val(.Cell(flexcpData, idX, cols.TipoDoc))
            
            If Trim(.Cell(flexcpText, idX, cols.Numero)) <> "" Then
                If Trim(arrDAT(0)) <> "" Then rsAdd!CDoSerie = Trim(arrDAT(0)) Else rsAdd!CDoSerie = Null
            Else
                 rsAdd!CDoSerie = Null
            End If
            
            rsAdd!CDoNumero = Null
            If UBound(arrDAT) > 0 Then If IsNumeric(arrDAT(1)) Then rsAdd!CDoNumero = Trim(arrDAT(1))
            
            rsAdd!CDoEstado = Val(.Cell(flexcpData, idX, cols.Estado))
            rsAdd("CDoUsuario") = paCodigoDeUsuario
            rsAdd.Update
        Next
    End With
    rsAdd.Close

    Screen.MousePointer = 0
    
    On Error Resume Next
    LimpioFicha
    Exit Sub

errGrabar:
    clsGeneral.OcurrioError "Error al grabar los datos.", Err.Description
    Screen.MousePointer = 0: Exit Sub
End Sub

Private Function fnc_Anulado(xEstado As Byte)
Dim x1 As Long, idX As Long
Dim bEXIT As Boolean
    dFila1 = fnc_TO_OBJ(vsLista.Row, xAnulado:=xEstado)
    If dFila1.idEstado <> 0 Then Exit Function
    
    dFila1.idEstado = xEstado
    If xEstado = 1 Or xEstado = 11 Then
        dFila1.idTipoDoc = 0
        dFila1.txtDocumento = ""
        dFila1.txtTipoDoc = ""
    End If

    If dFila1.idEstado = 10 Then
        fnc_FROM_OBJ dFila1, vsLista.Row
        Exit Function
    End If
    x1 = vsLista.Row
    For idX = x1 To vsLista.Rows - 1
        dFila2 = fnc_TO_OBJ(idX)
        
        fnc_FROM_OBJ dFila1, idX
        dFila1 = dFila2
    Next
    
    vsLista.AddItem ""
    If prmROJO_ASC Then
        vsLista.Cell(flexcpText, vsLista.Rows - 1, cols.Rojo) = vsLista.Cell(flexcpValue, vsLista.Rows - 2, cols.Rojo) + 1
    Else
        vsLista.Cell(flexcpText, vsLista.Rows - 1, cols.Rojo) = vsLista.Cell(flexcpValue, vsLista.Rows - 2, cols.Rojo) - 1
    End If
    fnc_FROM_OBJ dFila1, vsLista.Rows - 1
        
End Function

Private Function fnc_NoEsAnulado()
Dim x1 As Long, mROJO1 As String, mROJO2 As String
Dim idX As Long

    x1 = vsLista.Row
    dFila1 = fnc_TO_OBJ(x1)
    If dFila1.txtDocumento <> "" And dFila1.txtTipoDoc <> "" Then
        dFila1.idEstado = 0
        fnc_FROM_OBJ dFila1, x1
        Exit Function
    End If
    mROJO1 = vsLista.Cell(flexcpText, x1, cols.Rojo)
    vsLista.RemoveItem x1
    
    For idX = x1 To (vsLista.Rows - 1)
        mROJO2 = vsLista.Cell(flexcpText, idX, cols.Rojo)
        vsLista.Cell(flexcpText, idX, cols.Rojo) = mROJO1
        mROJO1 = mROJO2
    Next
    
End Function

Private Function fnc_CambiarF1xF2(xRow1 As Long, xRow2 As Long)
Dim x1 As Long, idX As Long

        dFila2 = fnc_TO_OBJ(xRow2)
        dFila1 = fnc_TO_OBJ(xRow1)
        fnc_FROM_OBJ dFila1, xRow2
        
        If (xRow2 - xRow1 - 1) > 0 Then
            x1 = xRow1
            For idX = x1 To xRow2 - 2
                dFila1 = fnc_TO_OBJ(idX + 1)
                fnc_FROM_OBJ dFila1, idX
            
            Next
            fnc_FROM_OBJ dFila2, idX
            
        Else
            fnc_FROM_OBJ dFila2, xRow1
        End If
        
End Function

Private Function fnc_TO_OBJ(xROW As Long, Optional xAnulado As Byte) As typData
    
    With fnc_TO_OBJ
        .txtTipoDoc = vsLista.Cell(flexcpText, xROW, cols.TipoDoc)
        .txtDocumento = vsLista.Cell(flexcpText, xROW, cols.Numero)
        
        .idEstado = Val(vsLista.Cell(flexcpData, xROW, cols.Estado))
        
        .idTipoDoc = vsLista.Cell(flexcpData, xROW, cols.TipoDoc)
    End With
    
End Function

Private Function fnc_FROM_OBJ(OBJX As typData, xROW As Long)
    With vsLista
        .Cell(flexcpData, xROW, cols.TipoDoc) = OBJX.idTipoDoc
        .Cell(flexcpText, xROW, cols.TipoDoc) = OBJX.txtTipoDoc
        .Cell(flexcpText, xROW, cols.Numero) = OBJX.txtDocumento
        
        .Cell(flexcpData, xROW, cols.Estado) = IIf(OBJX.idEstado <> 0, OBJX.idEstado, "")
        
        Select Case OBJX.idEstado
            Case 1: .Cell(flexcpText, xROW, cols.Estado) = "PA"
            Case 10: .Cell(flexcpText, xROW, cols.Estado) = "EXT"
            Case 11: .Cell(flexcpText, xROW, cols.Estado) = "ExA"
            Case Else: .Cell(flexcpText, xROW, cols.Estado) = ""
        End Select

    End With
End Function

Private Sub tFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtSerieRojo.SetFocus
End Sub

Private Sub tRojoD_GotFocus()
    tRojoD.SelStart = 0: tRojoD.SelLength = Len(tRojoD.Text)
End Sub

Private Sub tRojoD_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tRojoH.SetFocus
End Sub

Private Sub tRojoH_GotFocus()
    tRojoH.SelStart = 0: tRojoH.SelLength = Len(tRojoH.Text)
End Sub

Private Sub tRojoH_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        On Error Resume Next
        bCargar.SetFocus
    End If
    
End Sub


Private Sub txtSerieRojo_GotFocus()
    txtSerieRojo.SelStart = 0
    txtSerieRojo.SelLength = Len(txtSerieRojo.Text)
End Sub

Private Sub txtSerieRojo_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then tRojoD.SetFocus
End Sub

Private Sub vsLista_DblClick()
    On Error Resume Next
    If vsLista.Rows > vsLista.FixedRows Then
        If Val(bAccion(2).Tag) <> 0 Then Call bAccion_Click(Index:=2)
    End If
End Sub

Private Function fnc_CargoGrabados()

    mSQL = "Select * from ControlDocumentos" & _
                " Where CDoPreimpreso = " & prmIDPreimpreso & _
                " And CDoNumeroRojo Between  " & Val(tRojoD.Text) & " AND " & Val(tRojoH.Text) & _
                " AND CDoSerieRoja = '" & txtSerieRojo.Text & "'" & _
                " Order by CDoNumeroRojo"
    
    Dim mValor As Long
    Set RsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        With vsLista
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, cols.Rojo) = RsAux("CDoNumeroRojo").Value
            
            If Not IsNull(RsAux!CDoTipoDoc) Then
                .Cell(flexcpText, .Rows - 1, cols.TipoDoc) = RetornoNombreDocumento(RsAux!CDoTipoDoc, True)
                mValor = RsAux!CDoTipoDoc: .Cell(flexcpData, .Rows - 1, cols.TipoDoc) = mValor
            End If
            If Not IsNull(RsAux("CDoSerieRoja")) Then .Cell(flexcpData, .Rows - 1, 2) = Trim(RsAux("CDoSerieRoja"))
            
            If Not IsNull(RsAux!CDoSerie) And Not IsNull(RsAux!CDoNumero) Then .Cell(flexcpText, .Rows - 1, cols.Numero) = Trim(RsAux!CDoSerie) & "-" & RsAux!CDoNumero
            If Not IsNull(RsAux!CDoEstado) Then
                Select Case RsAux!CDoEstado
                    Case 1: .Cell(flexcpText, .Rows - 1, cols.Estado) = "PA"
                    Case 10: .Cell(flexcpText, .Rows - 1, cols.Estado) = "EXT"
                    Case 11: .Cell(flexcpText, .Rows - 1, cols.Estado) = "ExA"
                End Select
                mValor = RsAux!CDoEstado: .Cell(flexcpData, .Rows - 1, cols.Estado) = mValor
            End If
            
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close

End Function

Private Function fnc_EditarDatos()
On Error GoTo errEditar
Dim rowEdit As Long
Dim xCaso As Integer

    rowEdit = vsLista.Row
    xCaso = cTipo.ListIndex
    
    dFila1 = fnc_TO_OBJ(rowEdit)
    frmEdit.prm_Ed_Caso = xCaso
    
    frmEdit.prm_Ed_ROJO = vsLista.Cell(flexcpText, rowEdit, cols.Rojo)
    frmEdit.prm_Ed_TipoD = dFila1.idTipoDoc
    frmEdit.prm_Ed_Numero = dFila1.txtDocumento
    frmEdit.prm_Tipos = prmTipoDocs
5    frmEdit.prm_Ed_SerieRoja = vsLista.Cell(flexcpData, rowEdit, 2)
    
    frmEdit.prm_Ed_Estado = dFila1.idEstado
    
    frmEdit.Show vbModal, Me
    
    If frmEdit.prm_Grabar Then
    
        Dim arrDAT() As String
        arrDAT = Split(frmEdit.prm_Ed_Numero, "-")
        Dim mSERIE As String, mNumero As String
        If UBound(arrDAT) > 0 Then
            mSERIE = Trim(arrDAT(0))
            mNumero = Trim(arrDAT(1))
        ElseIf UBound(arrDAT) = 0 Then
            mNumero = Trim(arrDAT(0))
            mSERIE = ""
        Else
            mNumero = ""
            mSERIE = ""
        End If
            

        Dim bSalir As Boolean
        'Si cambio el NRO  Rojo  --Valido que no existaa
        If Not (frmEdit.prm_Ed_ROJO = vsLista.Cell(flexcpText, rowEdit, cols.Rojo)) Then
        
            mSQL = "Select * from ControlDocumentos Where CDoPreimpreso = " & prmIDPreimpreso & _
                        " And CDoNumeroRojo = " & frmEdit.prm_Ed_ROJO _
                        & " AND CDoSerieRoja = '" & frmEdit.prm_Ed_SerieRoja & "'"

            Set RsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then bSalir = True
            RsAux.Close
            
            If bSalir Then
                MsgBox "El número ROJO ingresado ya fue asignado !!", vbExclamation, "Duplicación"
                Exit Function
            End If
        End If
        
        If mNumero <> "" Then
            If Not (frmEdit.prm_Ed_TipoD = dFila1.idTipoDoc And LCase(frmEdit.prm_Ed_Numero) = LCase(dFila1.txtDocumento)) Then
                                
                mSQL = "Select * from ControlDocumentos " & _
                            " Where CDoPreimpreso = " & prmIDPreimpreso & _
                            " And CDoNumero = " & mNumero & " AND CDoSerieRoja = '" & frmEdit.prm_Ed_SerieRoja & "'" & _
                            " And CDoTipoDoc = " & frmEdit.prm_Ed_TipoD & _
                            " And CDoEstado IN (0, 10)" & _
                            IIf(mSERIE <> "", " And CDoSerie = '" & mSERIE & "'", "")
                
                Set RsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
                If Not RsAux.EOF Then bSalir = True
                RsAux.Close
        
                If bSalir Then
                    MsgBox "El documento ingresado ya fue asignado a otro número rojo !!", vbExclamation, "Duplicación"
                    Exit Function
                End If
                
            End If
        End If
        
        'EDITO LOS DATOS    ------------------------------------------------------------------------------
        mSQL = "Select * from ControlDocumentos " & _
                    " Where CDoPreimpreso = " & prmIDPreimpreso & _
                    " And CDoNumeroRojo = " & vsLista.Cell(flexcpText, rowEdit, cols.Rojo) _
                    & " AND CDoSerieRoja = '" & frmEdit.prm_Ed_SerieRoja & "'"
        Set RsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
        
        RsAux.Edit
    
        RsAux!CDONumeroRojo = frmEdit.prm_Ed_ROJO
        RsAux("CDoSerieRoja") = frmEdit.prm_Ed_SerieRoja
        RsAux!CDoTipoDoc = frmEdit.prm_Ed_TipoD
        
        If mSERIE <> "" Then RsAux!CDoSerie = mSERIE Else RsAux!CDoSerie = Null
        If mNumero <> "" Then RsAux!CDoNumero = mNumero Else RsAux!CDoNumero = Null
        If xCaso = 1 Then
            RsAux!CDoEstado = frmEdit.prm_Ed_Estado
            If frmEdit.prm_Ed_Estado = 1 Then  'PA
                RsAux!CDoSerie = Null
                RsAux!CDoNumero = Null
                RsAux!CDoTipoDoc = Null
            End If
        End If
        
        RsAux("CDoUsuario") = paCodigoDeUsuario
            
        RsAux.Update
        RsAux.Close
        
        With vsLista
            .Cell(flexcpText, rowEdit, cols.Rojo) = frmEdit.prm_Ed_ROJO
            
            .Cell(flexcpText, rowEdit, cols.TipoDoc) = RetornoNombreDocumento(frmEdit.prm_Ed_TipoD, True)
            .Cell(flexcpData, rowEdit, cols.TipoDoc) = frmEdit.prm_Ed_TipoD
            .Cell(flexcpData, rowEdit, 2) = frmEdit.prm_Ed_SerieRoja
            
            
            If mSERIE <> "" And mNumero <> "" Then .Cell(flexcpText, rowEdit, cols.Numero) = mSERIE & "-" & mNumero
        End With
        If xCaso = 1 Then
            'fnc_Anulado xEstado:=frmEdit.prm_Ed_Estado
            dFila1 = fnc_TO_OBJ(rowEdit, xAnulado:=frmEdit.prm_Ed_Estado)
            dFila1.idEstado = frmEdit.prm_Ed_Estado
            If dFila1.idEstado = 1 Or dFila1.idEstado = 11 Then
                dFila1.idTipoDoc = 0
                dFila1.txtDocumento = ""
                dFila1.txtTipoDoc = ""
            End If
            
            fnc_FROM_OBJ dFila1, vsLista.Row
        End If
            
            
    End If
    Exit Function
    
errEditar:
    clsGeneral.OcurrioError "Error al editar los datos..", Err.Number & "- " & Err.Description
End Function
