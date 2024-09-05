VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmListado 
   Caption         =   "Corregir Existencia"
   ClientHeight    =   7530
   ClientLeft      =   1170
   ClientTop       =   1905
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   11880
   Begin VB.TextBox tEQ 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5820
      TabIndex        =   14
      Top             =   780
      Width           =   675
   End
   Begin VB.TextBox tEArticulo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   780
      Width           =   4215
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   120
      TabIndex        =   6
      Top             =   1140
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7646
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
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   12
      FixedRows       =   1
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
      OutlineBar      =   1
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
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   7275
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   12753
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fFiltros 
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   11415
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   9720
         Picture         =   "frmListado.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   240
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   10560
         Picture         =   "frmListado.frx":0744
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   240
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   10080
         Picture         =   "frmListado.frx":0846
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   240
         Width           =   310
      End
      Begin VB.TextBox tFHasta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   3
         Text            =   "28/12/2000"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox tArticulo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5400
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Text            =   "28/12/2000"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Hasta:"
         Height          =   255
         Left            =   2880
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   4680
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha de Compra:"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   255
         Width           =   1455
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Q:"
      Height          =   255
      Left            =   5520
      TabIndex        =   15
      Top             =   780
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ar&tículo:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   780
      Width           =   735
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum TipoCV
    Compra = 1
    Comercio = 2
    Importacion = 3
End Enum

Private RsAux As rdoResultset, Rs1 As rdoResultset
Private aTexto As String

Private Sub AccionLimpiar()
    tFecha.Text = "": tFHasta.Text = ""
    tArticulo.Text = "": tArticulo.Tag = 0
    vsConsulta.Rows = 1
    EditoRenglon False
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub
Private Sub bNoFiltros_Click()
    AccionLimpiar
End Sub

Private Sub Label1_Click()
    Foco tArticulo
End Sub

Private Sub Label2_Click()
    Foco tFHasta
End Sub

Private Sub Label3_Click()
    Foco tEArticulo
End Sub

Private Sub Label4_Click()
    Foco tEQ
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = "0"
End Sub

Private Sub tArticulo_GotFocus()
    With tArticulo: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tArticulo_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo ErrTA
    
    If KeyCode = vbKeyReturn And Trim(tArticulo.Text) <> "" Then
        tArticulo.Tag = "0"
        Screen.MousePointer = 11
        If Not IsNumeric(tArticulo.Text) Then
            Cons = "Select ArtID, 'Código' = ArtCodigo, 'Nombre' = ArtNombre From Articulo Where ArtNombre Like '" & tArticulo.Text & "%'"
            Dim LiAyuda  As New clsListadeAyuda
            LiAyuda.ActivoListaAyuda Cons, False, miConexion.TextoConexion(logComercio), 4300
            If LiAyuda.ItemSeleccionado <> "" Then tArticulo.Text = LiAyuda.ItemSeleccionado Else tArticulo.Text = "0"
            Set LiAyuda = Nothing
        End If
        
        If CLng(tArticulo.Text) > 0 Then
            Cons = "Select ArtID, 'Código' = ArtCodigo, 'Nombre' = ArtNombre From Articulo Where ArtCodigo = " & CLng(tArticulo.Text)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
            If RsAux.EOF Then
                RsAux.Close
                MsgBox "No se encontró un artículo con ese código.", vbInformation, "ATENCIÓN"
            Else
                tArticulo.Text = Trim(RsAux!Nombre)
                tArticulo.Tag = RsAux!ArtId
                RsAux.Close
                Foco bConsultar
            End If
        Else
            tArticulo.Text = "0"
        End If
        Screen.MousePointer = 0
    Else
        If KeyCode = vbKeyReturn Then Foco bConsultar
    End If
    Exit Sub
ErrTA:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tEArticulo_Change()
    tEArticulo.Tag = 0
End Sub

Private Sub tEArticulo_GotFocus()
    With tEArticulo: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tEArticulo_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo ErrTA
    If KeyCode = vbKeyEscape Then EditoRenglon False: Exit Sub
    
    If KeyCode = vbKeyReturn And Trim(tEArticulo.Text) <> "" Then
        If Val(tEArticulo.Tag) > 0 Then Foco tEQ: Exit Sub
        
        Screen.MousePointer = 11
        If Not IsNumeric(tEArticulo.Text) Then
            Cons = "Select ArtID, 'Nombre' = ArtNombre,  'Código' = ArtCodigo From Articulo Where ArtNombre Like '" & tEArticulo.Text & "%'"
            Dim LiAyuda  As New clsListadeAyuda
            LiAyuda.ActivoListaAyuda Cons, False, miConexion.TextoConexion(logComercio), 4300
            If LiAyuda.ItemSeleccionado <> "" Then
                tEArticulo.Text = LiAyuda.ItemSeleccionado
                tEArticulo.Tag = LiAyuda.ValorSeleccionado
            End If
            Set LiAyuda = Nothing
        
        Else
            Cons = "Select ArtID, 'Código' = ArtCodigo, 'Nombre' = ArtNombre From Articulo Where ArtCodigo = " & CLng(tEArticulo.Text)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
            If RsAux.EOF Then
                RsAux.Close
                MsgBox "No se encontró un artículo con ese código.", vbInformation, "ATENCIÓN"
            Else
                tEArticulo.Text = Trim(RsAux!Nombre)
                tEArticulo.Tag = RsAux!ArtId
                RsAux.Close
                Foco tEQ
            End If
        End If
        
        Screen.MousePointer = 0
    Else
        If KeyCode = vbKeyReturn Then Foco tEQ
    End If
    Exit Sub
ErrTA:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub tEQ_GotFocus()
    With tEQ: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tEQ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then EditoRenglon False: Exit Sub
End Sub

Private Sub tEQ_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
    
        If Val(tEArticulo.Tag) = 0 Then
            MsgBox "Debe ingresar el artículo que figura en la existencia.", vbExclamation, "ATENCIÓN"
            Foco tEArticulo: Exit Sub
        End If
        
        If Not IsNumeric(tEQ.Text) Then
            MsgBox "La cantidad ingresada no es correcta. Para eliminar la existencia ingrese valor cero.", vbExclamation, "ATENCIÓN"
            Foco tEQ: Exit Sub
        End If
        
        If MsgBox("Confirma actualizar la existencia con los siguientes valores:" & Chr(vbKeyReturn) _
                    & "Artículo: " & Trim(tEArticulo.Text) & Chr(vbKeyReturn) _
                    & "Cantidad: " & Trim(tEQ.Text), vbQuestion + vbYesNo, "Actualizar existencia") = vbNo Then Exit Sub
        
        AccionGrabar
    End If
    
End Sub

Private Sub AccionGrabar()

    'Data -->   0)Id_articulo     2)Id_Compra     3)ComTipo
    'Tag de vscosnulta  = idRow
    On Error GoTo errGrabar
    Dim Row As Long: Row = Val(vsConsulta.Tag)
    If Row = 0 Then Exit Sub
    Screen.MousePointer = 11
    aTexto = ""
    With vsConsulta
        Cons = "Select * from CMCompra " _
                & " Where ComFecha = '" & Format(.Cell(flexcpText, Row, 3), sqlFormatoF) & "'" _
                & " And ComArticulo = " & .Cell(flexcpData, Row, 0) _
                & " And ComCodigo = " & .Cell(flexcpData, Row, 2) _
                & " And ComTipo = " & .Cell(flexcpData, Row, 3)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            If CCur(tEQ.Text) = 0 Then
                RsAux.Delete
                aTexto = "El registro de existencia ha sido eliminado."
                .RemoveItem Row
            Else
                RsAux.Edit
                RsAux!ComArticulo = Val(tEArticulo.Tag)
                RsAux!ComCantidad = CCur(tEQ.Text)
                RsAux.Update
                RsAux.Close
                
                'Corrigo los datos en la grilla--------------------------------------------------------------------------------------------
                Cons = "Select * from CMCompra, Articulo " _
                        & " Where ComFecha = '" & Format(.Cell(flexcpText, Row, 3), sqlFormatoF) & "'" _
                        & " And ComArticulo = " & Val(tEArticulo.Tag) _
                        & " And ComCodigo = " & .Cell(flexcpData, Row, 2) _
                        & " And ComTipo = " & .Cell(flexcpData, Row, 3) _
                        & " And ComArticulo = ArtID"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                
                'Data -->   0)Id_articulo     2)Id_Compra     3)ComTipo
                Dim aValor As Long
                .Cell(flexcpText, Row, 1) = Format(RsAux!ArtCodigo, "(#,000,000)") & " " & Trim(RsAux!ArtNombre)
                aValor = RsAux!ComArticulo: .Cell(flexcpData, Row, 0) = aValor
                
                If RsAux!ComTipo = TipoCV.Compra Then .Cell(flexcpText, Row, 2) = RsAux!ComCodigo
                aValor = RsAux!ComCodigo: .Cell(flexcpData, Row, 2) = aValor
                
                .Cell(flexcpText, Row, 3) = Format(RsAux!ComFecha, "dd/mm/yy")
                aValor = RsAux!ComTipo: .Cell(flexcpData, Row, 3) = aValor
                
                .Cell(flexcpText, Row, 4) = RsAux!ComCantidad
                .Cell(flexcpText, Row, 5) = Format(RsAux!ComCosto, FormatoMonedaP)
                .Cell(flexcpText, Row, 6) = Format(RsAux!ComCantidad * RsAux!ComCosto, FormatoMonedaP)
                If RsAux!ComTipo = TipoCV.Compra And .Cell(flexcpValue, Row, 7) <> 0 Then
                    .Cell(flexcpText, Row, 8) = Format(CCur(.Cell(flexcpText, Row, 6)) / .Cell(flexcpValue, Row, 7), FormatoMonedaP)
                End If
                '----------------------------------------------------------------------------------------------------------------------------
                aTexto = "El registro de la existencia ha sido modificado."
            End If
        End If
        RsAux.Close
    End With
    If Trim(aTexto) = "" Then
        aTexto = "Ocurrió un error al modificar el registro de existencia."
    Else
        EditoRenglon False
    End If
    MsgBox aTexto, vbInformation, "Información"
    Screen.MousePointer = 0
    Exit Sub

errGrabar:
    clsGeneral.OcurrioError "Ocurrió un error al modificar el registro.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tFecha_GotFocus()
    With tFecha: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tFHasta
End Sub
Private Sub tFecha_LostFocus()
    If IsDate(tFecha.Text) Then
        tFecha.Text = Format(tFecha.Text, FormatoFP)
        If Not IsDate(tFHasta.Text) Then tFHasta.Text = tFecha.Text
    End If
End Sub

Private Sub Label5_Click()
    Foco tFecha
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub

Private Sub Form_Load()
    On Error GoTo ErrLoad
    
    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    InicializoGrillas
    AccionLimpiar
    EditoRenglon False
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsConsulta
        .OutlineBar = flexOutlineBarNone ' = flexOutlineBarComplete
        .OutlineCol = 0
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        
        .Cols = 1: .Rows = 1:
        .FormatString = "<|<Artículo|<Compra|<Fecha|>Q|>Costo|>Total|>TC|>U$S|"
            
        .WordWrap = False
        .ColWidth(0) = 0: .ColWidth(1) = 2700: .ColWidth(2) = 650: .ColWidth(3) = 800: .ColWidth(4) = 1100
        .ColWidth(5) = 1400: .ColWidth(6) = 1400: .ColWidth(7) = 600: .ColWidth(8) = 1300: .ColWidth(9) = 10

        .MergeCol(0) = True: .ColAlignment(0) = flexAlignLeftBottom
        .MergeCells = flexMergeSpill
    End With
      
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyE: AccionConsultar
            Case vbKeyX: Unload Me
        End Select
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Screen.MousePointer = 11

    vsConsulta.Height = Me.ScaleHeight - (vsConsulta.Top + Status.Height + 70)
    
    fFiltros.Width = Me.ScaleWidth - (vsConsulta.Left * 2)
    vsConsulta.Width = fFiltros.Width
    vsConsulta.Left = fFiltros.Left
    
    Screen.MousePointer = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    GuardoSeteoForm Me
    
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
    
End Sub

Private Sub AccionConsultar()
    
    EditoRenglon False
    
    If Not IsDate(tFecha.Text) And IsDate(tFHasta.Text) Then
        MsgBox "Ingrese la fecha desde.", vbExclamation, "ATENCIÓN"
        Foco tFecha: Exit Sub
    End If
    If IsDate(tFecha.Text) And Not IsDate(tFHasta.Text) Then
        If Trim(tFHasta.Text) = "" Then
            tFHasta.Text = tFecha.Text
        Else
            MsgBox "La fecha hasta no es correcta.", vbExclamation, "ATENCIÓN"
            Foco tFHasta: Exit Sub
        End If
    End If
    If IsDate(tFecha.Text) And IsDate(tFHasta.Text) Then
        If CDate(tFecha.Text) > CDate(tFHasta.Text) Then
            MsgBox "Los rangos de fecha no son correctos.", vbExclamation, "ATENCIÓN"
            Foco tFecha: Exit Sub
        End If
    End If
    On Error GoTo errConsultar
    Screen.MousePointer = 11
   
    
    Cons = "Select ArtID, ArtCodigo, ArtNombre, 'IDCompra' = Compra.ComCodigo, 'Fecha' = CM.ComFecha, CM.ComCodigo, ComCantidad, ComCosto, CM.ComTipo, ComTC " _
        & " From Articulo, CMCompra CM" _
                    & " Left Outer Join Compra On Compra.ComCodigo = CM.ComCodigo" _
            & " Where ComArticulo = ArtID "
    If IsDate(tFecha.Text) Then
        Cons = Cons & " And CM.ComFecha >= '" & Format(tFecha.Text & " 00:00:00", sqlFormatoFH) & "'" _
                           & " And CM.ComFecha <= '" & Format(tFHasta.Text & " 23:59:59", sqlFormatoFH) & "'"
    End If
    If Val(tArticulo.Tag) > 0 Then Cons = Cons & " And ComArticulo = " & CLng(tArticulo.Tag)
    Cons = Cons & " Order by ArtNombre, Fecha DESC"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        RsAux.Close: Screen.MousePointer = 0: InicializoGrillas: Exit Sub
    End If
    Dim aIdAnterior As Long: aIdAnterior = 0
    Dim aValor As Long
    With vsConsulta
        .Rows = 1
        Do While Not RsAux.EOF
            'Data -->   0)Id_articulo     2)Id_Compra     3)ComTipo
                        
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!ArtCodigo, "(#,000,000)") & " " & Trim(RsAux!ArtNombre)
            aValor = RsAux!ArtId: .Cell(flexcpData, .Rows - 1, 0) = aValor
            
            If RsAux!ComTipo = TipoCV.Compra Then .Cell(flexcpText, .Rows - 1, 2) = RsAux!IdCompra
            aValor = RsAux!ComCodigo: .Cell(flexcpData, .Rows - 1, 2) = aValor
            
            .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!fecha, "dd/mm/yy")
            aValor = RsAux!ComTipo: .Cell(flexcpData, .Rows - 1, 3) = aValor
            
            .Cell(flexcpText, .Rows - 1, 4) = RsAux!ComCantidad
            .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!ComCosto, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!ComCantidad * RsAux!ComCosto, FormatoMonedaP)
            If RsAux!ComTipo = TipoCV.Compra Then
                .Cell(flexcpText, .Rows - 1, 7) = Format(RsAux!ComTC, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 8) = Format(CCur(.Cell(flexcpText, .Rows - 1, 6)) / RsAux!ComTC, FormatoMonedaP)
            End If
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        .Select 1, 0, 1, 1
        .Sort = flexSortGenericAscending
             
    End With
    
    Screen.MousePointer = 0
    Exit Sub
errConsultar:
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tFHasta_GotFocus()
    With tFHasta: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tFHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tArticulo
End Sub

Private Sub tFHasta_LostFocus()
    If IsDate(tFHasta.Text) Then tFHasta.Text = Format(tFHasta.Text, FormatoFP)
End Sub

Private Sub EditoRenglon(Estado As Boolean, Optional Row As Long = 0)
    On Error Resume Next
    tEArticulo.Enabled = Estado: tEQ.Enabled = Estado
    
    If Estado Then
        tEArticulo.BackColor = Colores.Blanco
        tEQ.BackColor = Colores.Blanco
        tEArticulo.Text = vsConsulta.Cell(flexcpText, Row, 1): tEArticulo.Tag = vsConsulta.Cell(flexcpData, Row, 0)
        tEQ.Text = vsConsulta.Cell(flexcpText, Row, 4)
        vsConsulta.Enabled = False: vsConsulta.BackColor = Colores.Inactivo
        Foco tEArticulo
    Else
        tEArticulo.BackColor = Colores.Inactivo
        tEQ.BackColor = Colores.Inactivo
        tEArticulo.Text = "": tEQ.Text = ""
        vsConsulta.Enabled = True: vsConsulta.BackColor = Colores.Blanco
        vsConsulta.SetFocus
    End If
    
    vsConsulta.Tag = Row
    
End Sub

Private Sub vsConsulta_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And vsConsulta.Rows > 1 Then EditoRenglon True, vsConsulta.Row
    
End Sub

