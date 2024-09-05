VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form frmListado 
   Caption         =   "Corregir Existencia"
   ClientHeight    =   6885
   ClientLeft      =   5985
   ClientTop       =   5970
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
   ScaleHeight     =   6885
   ScaleWidth      =   11880
   Begin AACombo99.AACombo cComentario 
      Height          =   315
      Left            =   2040
      TabIndex        =   19
      Top             =   1500
      Width           =   7755
      _ExtentX        =   13679
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
   Begin VB.CommandButton bNoFiltros 
      Height          =   310
      Left            =   480
      Picture         =   "frmListado.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Quitar filtros."
      Top             =   1140
      Width           =   310
   End
   Begin VB.CommandButton bConsultar 
      Height          =   310
      Left            =   120
      Picture         =   "frmListado.frx":0808
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Ejecutar."
      Top             =   1140
      Width           =   310
   End
   Begin VB.CommandButton bNuevo 
      Height          =   310
      Left            =   900
      Picture         =   "frmListado.frx":0B0A
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Ejecutar."
      Top             =   1140
      Width           =   310
   End
   Begin VB.TextBox tECosto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8820
      TabIndex        =   17
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox tEFecha 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7260
      TabIndex        =   15
      Top             =   1200
      Width           =   915
   End
   Begin VB.TextBox tEQ 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5940
      TabIndex        =   13
      Top             =   1200
      Width           =   675
   End
   Begin VB.TextBox tEArticulo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1200
      Width           =   3555
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4095
      Left            =   120
      TabIndex        =   8
      Top             =   1860
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7223
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
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   6630
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "terminal"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "bd"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12753
         EndProperty
      EndProperty
   End
   Begin VB.Frame fFiltros 
      Caption         =   "Filtros para Consultar "
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
      Height          =   1035
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   11415
      Begin VB.CommandButton bAddMultiple 
         Caption         =   "RR"
         Height          =   315
         Left            =   9780
         TabIndex        =   26
         ToolTipText     =   "Resumen de Rebotes"
         Top             =   600
         Width           =   435
      End
      Begin AACombo99.AACombo cGrupos 
         Height          =   315
         Left            =   1680
         TabIndex        =   24
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
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
      Begin VB.TextBox tNombres 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5400
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   600
         Width           =   4215
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
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "&Grupos de Artículos:"
         Height          =   255
         Left            =   180
         TabIndex        =   25
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nombres:"
         Height          =   255
         Left            =   4680
         TabIndex        =   6
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Hasta:"
         Height          =   255
         Left            =   2880
         TabIndex        =   2
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   4680
         TabIndex        =   4
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha de Compra:"
         Height          =   255
         Left            =   180
         TabIndex        =   0
         Top             =   315
         Width           =   1455
      End
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "&Coms..:"
      Height          =   255
      Left            =   1320
      TabIndex        =   18
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Cos&to:"
      Height          =   255
      Left            =   8220
      TabIndex        =   16
      Top             =   1200
      Width           =   555
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Fe&cha:"
      Height          =   255
      Left            =   6720
      TabIndex        =   14
      Top             =   1200
      Width           =   555
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Q:"
      Height          =   255
      Left            =   5700
      TabIndex        =   12
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ar&tículo:"
      Height          =   255
      Left            =   1320
      TabIndex        =   20
      Top             =   1200
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

Private rsAux As rdoResultset, rs1 As rdoResultset
Private aTexto As String
Dim sNuevo As Boolean


Private Sub AccionLimpiar()
On Error Resume Next
    tFecha.Text = "": tFHasta.Text = ""
    tArticulo.Text = "": tArticulo.Tag = 0
    
    tNombres.Text = ""
    cGrupos.Text = ""
    
    vsConsulta.Rows = 1
    EditoRenglon False
    
End Sub

Private Sub bAddMultiple_Click()
    frmMultiple.Show vbModal
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub
Private Sub bNoFiltros_Click()
    AccionLimpiar
End Sub

Private Sub bNuevo_Click()
    EditoRenglonParaNuevo True
End Sub

Private Sub cComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub cGrupos_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tNombres
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
        
            Dim LiAyuda  As New clsListadeAyuda
            
            tArticulo.Text = Replace(Trim(tArticulo.Text), " ", "%")
            cons = "Select ArtID, 'Código' = ArtCodigo, 'Nombre' = ArtNombre From Articulo " & _
                       " Where ArtNombre Like '" & tArticulo.Text & "%'"
            
            If LiAyuda.ActivarAyuda(cBase, cons, 4300, 1, "Lista de Artículos") <> 0 Then
                tArticulo.Text = LiAyuda.RetornoDatoSeleccionado(1)
            Else
                tArticulo.Text = "0"
            End If
            Set LiAyuda = Nothing
        End If
        
        If CLng(tArticulo.Text) > 0 Then
            cons = "Select ArtID, 'Código' = ArtCodigo, 'Nombre' = ArtNombre From Articulo Where ArtCodigo = " & CLng(tArticulo.Text)
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurReadOnly)
            If rsAux.EOF Then
                rsAux.Close
                MsgBox "No se encontró un artículo con ese código.", vbInformation, "ATENCIÓN"
            Else
                tArticulo.Text = Format(rsAux(1), "(#,000,000)") & " " & Trim(rsAux!Nombre)
                tArticulo.Tag = rsAux!ArtID
                rsAux.Close
                Foco cGrupos
            End If
        Else
            tArticulo.Text = "0"
        End If
        Screen.MousePointer = 0
    Else
        If KeyCode = vbKeyReturn Then Foco cGrupos
    End If
    Exit Sub
ErrTA:
    clsGeneral.OcurrioError "Error al buscar el artículo.", Err.Description
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
            Dim LiAyuda  As New clsListadeAyuda
        
            tEArticulo.Text = Replace(Trim(tEArticulo.Text), " ", "%")
            cons = "Select ArtID, 'Nombre' = ArtNombre,  'Código' = ArtCodigo From Articulo Where ArtNombre Like '" & tEArticulo.Text & "%'"
            
            If LiAyuda.ActivarAyuda(cBase, cons, 4300, 1, "Lista de Artículos") <> 0 Then
                tEArticulo.Text = LiAyuda.RetornoDatoSeleccionado(1)
                tEArticulo.Tag = LiAyuda.RetornoDatoSeleccionado(0)
            End If
            Set LiAyuda = Nothing
        
        Else
            cons = "Select ArtID, 'Código' = ArtCodigo, 'Nombre' = ArtNombre From Articulo Where ArtCodigo = " & CLng(tEArticulo.Text)
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurReadOnly)
            If rsAux.EOF Then
                rsAux.Close
                MsgBox "No se encontró un artículo con ese código.", vbInformation, "ATENCIÓN"
            Else
                tEArticulo.Text = Trim(rsAux!Nombre)
                tEArticulo.Tag = rsAux!ArtID
                rsAux.Close
                Foco tEQ
            End If
        End If
        
        Screen.MousePointer = 0
    Else
        If KeyCode = vbKeyReturn Then Foco tEQ
    End If
    Exit Sub
ErrTA:
    clsGeneral.OcurrioError "Error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub tECosto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then EditoRenglon False: Exit Sub
End Sub

Private Sub tECosto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cComentario.SetFocus
End Sub

Private Sub tECosto_LostFocus()
    If IsNumeric(tECosto.Text) Then tECosto.Text = Format(tECosto.Text, FormatoMonedaP)
End Sub

Private Sub tEFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then EditoRenglon False: Exit Sub
End Sub

Private Sub tEFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tECosto
End Sub

Private Sub tEFecha_LostFocus()
    If IsDate(tEFecha.Text) Then tEFecha.Text = Format(tEFecha.Text, "dd/mm/yyyy")
End Sub

Private Sub tEQ_GotFocus()
    With tEQ: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tEQ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then EditoRenglon False: Exit Sub
End Sub

Private Sub tEQ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tEFecha
End Sub

Private Function ValidoCampos() As Boolean
    
    ValidoCampos = False
    
    If Val(tEArticulo.Tag) = 0 Then
        MsgBox "Debe ingresar el artículo que figura en la existencia.", vbExclamation, "ATENCIÓN"
        Foco tEArticulo: Exit Function
    End If
    
    If Not IsNumeric(tEQ.Text) Then
        MsgBox "La cantidad ingresada no es correcta. Para eliminar la existencia ingrese valor cero.", vbExclamation, "ATENCIÓN"
        Foco tEQ: Exit Function
    End If
    
    If Not IsDate(tEFecha.Text) Then
        MsgBox "La fecha ingresada no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tEFecha: Exit Function
    End If
    
    If Not IsNumeric(tECosto.Text) Then
        MsgBox "El costo ingresado no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tECosto: Exit Function
    End If
    
    ValidoCampos = True
    
End Function

Private Sub AccionGrabar()
        
    If Not sNuevo Then
    
        If MsgBox("Confirma actualizar la existencia con los siguientes valores:" & Chr(vbKeyReturn) _
                    & "Artículo: " & Trim(tEArticulo.Text) & Chr(vbKeyReturn) _
                    & "Cantidad: " & Trim(tEQ.Text), vbQuestion + vbYesNo, "Actualizar existencia") = vbNo Then Exit Sub
        
        ModificoRegistroCompra
    
    Else
        
        If MsgBox("Confirma agregar a la existencia los siguientes valores:" & Chr(vbKeyReturn) _
                    & "Artículo: " & Trim(tEArticulo.Text) & Chr(vbKeyReturn) _
                    & "Cantidad: " & Trim(tEQ.Text), vbQuestion + vbYesNo, "Agregar existencia") = vbNo Then Exit Sub
        
        AgregoRegistroCMCompra
        
    End If
    
    
End Sub


Private Sub ModificoRegistroCompra()

    'Data -->   0)Id_articulo     2)Id_Compra     3)ComTipo
    'Tag de vscosnulta  = idRow
    On Error GoTo errGrabar
    Dim Row As Long: Row = Val(vsConsulta.Tag)
    
    If Row = 0 Then Exit Sub
    
    Screen.MousePointer = 11
    aTexto = ""
    With vsConsulta
        cons = "Select * from CMCompra " _
                & " Where ComFecha = '" & Format(.Cell(flexcpText, Row, 3), sqlFormatoF) & "'" _
                & " And ComArticulo = " & .Cell(flexcpData, Row, 0) _
                & " And ComCodigo = " & .Cell(flexcpData, Row, 2) _
                & " And ComTipo = " & .Cell(flexcpData, Row, 3)
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            If CCur(tEQ.Text) = 0 Then
                rsAux.Delete
                
                AgregoAjuste False, vsConsulta.Cell(flexcpData, Row, 3), vsConsulta.Cell(flexcpData, Row, 2), Row
                
                aTexto = "El registro de existencia ha sido eliminado."
                .RemoveItem Row
            Else
                rsAux.Edit
                rsAux!ComArticulo = Val(tEArticulo.Tag)
                rsAux!ComCantidad = CCur(tEQ.Text)
                rsAux!ComFecha = Format(tEFecha.Text, sqlFormatoF)
                rsAux!ComCosto = CCur(tECosto.Text)
                rsAux.Update
                rsAux.Close
                
                AgregoAjuste False, vsConsulta.Cell(flexcpData, Row, 3), vsConsulta.Cell(flexcpData, Row, 2), Row
                
                'Corrigo los datos en la grilla--------------------------------------------------------------------------------------------
                cons = "Select * from CMCompra, Articulo " _
                        & " Where ComFecha = '" & Format(.Cell(flexcpText, Row, 3), sqlFormatoF) & "'" _
                        & " And ComArticulo = " & Val(tEArticulo.Tag) _
                        & " And ComCodigo = " & .Cell(flexcpData, Row, 2) _
                        & " And ComTipo = " & .Cell(flexcpData, Row, 3) _
                        & " And ComArticulo = ArtID"
                Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                
                'Data -->   0)Id_articulo     2)Id_Compra     3)ComTipo
                Dim aValor As Long
                .Cell(flexcpText, Row, 1) = Format(rsAux!ArtCodigo, "(#,000,000)") & " " & Trim(rsAux!ArtNombre)
                aValor = rsAux!ComArticulo: .Cell(flexcpData, Row, 0) = aValor
                
                If rsAux!ComTipo = TipoCV.Compra Then .Cell(flexcpText, Row, 2) = rsAux!ComCodigo
                aValor = rsAux!ComCodigo: .Cell(flexcpData, Row, 2) = aValor
                
                .Cell(flexcpText, Row, 3) = Format(rsAux!ComFecha, "dd/mm/yy")
                aValor = rsAux!ComTipo: .Cell(flexcpData, Row, 3) = aValor
                
                .Cell(flexcpText, Row, 4) = rsAux!ComCantidad
                .Cell(flexcpText, Row, 5) = Format(rsAux!ComCosto, FormatoMonedaP)
                .Cell(flexcpText, Row, 6) = Format(rsAux!ComCantidad * rsAux!ComCosto, FormatoMonedaP)
                If rsAux!ComTipo = TipoCV.Compra And .Cell(flexcpValue, Row, 7) <> 0 Then
                    .Cell(flexcpText, Row, 8) = Format(CCur(.Cell(flexcpText, Row, 6)) / .Cell(flexcpValue, Row, 7), FormatoMonedaP)
                End If
                '----------------------------------------------------------------------------------------------------------------------------
                aTexto = "El registro de la existencia ha sido modificado."
            End If
        End If
        rsAux.Close
    End With
    
    AgregoComentario
    
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

Private Sub AgregoRegistroCMCompra()
    
    'Data -->   0)Id_articulo     2)Id_Compra     3)ComTipo
    'Tag de vscosnulta  = idRow
    
    On Error GoTo errGrabar
    Screen.MousePointer = 11
    Dim aMin As Long, bOK As Boolean
    
    aMin = 0: bOK = False
    cons = "Select Min(ComCodigo) From CMCompra"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then aMin = rsAux(0)
    rsAux.Close
    
    aMin = aMin - 1
    
    cons = "Select * from CMCompra " _
            & " Where ComFecha = '" & Format(tEFecha.Text, sqlFormatoF) & "'" _
            & " And ComArticulo = " & tEArticulo.Tag _
            & " And ComCodigo = " & aMin _
            & " And ComTipo = " & TipoCV.Comercio
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If rsAux.EOF Then
        rsAux.AddNew
        
        rsAux!ComFecha = Format(tEFecha.Text, sqlFormatoF)
        rsAux!ComArticulo = Val(tEArticulo.Tag)
        rsAux!ComCantidad = CCur(tEQ.Text)
        rsAux!ComCodigo = aMin
        rsAux!ComTipo = TipoCV.Comercio
        rsAux!ComCosto = CCur(tECosto.Text)
        rsAux!ComQOriginal = CCur(tEQ.Text)
        rsAux.Update
        bOK = True
    Else
        MsgBox "No se pudo grabar el nuevo registro (existen compras para el mismo id).", vbCritical, "ERROR"
    End If
    rsAux.Close
    
    If bOK Then AgregoAjuste True, TipoCV.Comercio, aMin, 0

    
    AgregoComentario
    tEArticulo.Text = "": tEQ.Text = ""
    Foco tEArticulo
    Screen.MousePointer = 0
    
    Exit Sub

errGrabar:
    clsGeneral.OcurrioError "Error al modificar el registro.", Err.Description
    Screen.MousePointer = 0
End Sub
 
Private Function AgregoAjuste(CasoNuevo As Boolean, mTipoCV As Integer, mCodigo As Long, Row As Long)
On Error GoTo errAjuste
Dim rsAj As rdoResultset
Dim pdatCambio As Date

    pdatCambio = Now

Dim pdtmFecha_A As Date, plngArticulo_A As Long, pcurQ_A As Long, pdblCosto_A As Double
Dim pdtmFecha_D As Date, plngArticulo_D As Long, pcurQ_D As Long, pdblCosto_D As Double


    plngArticulo_D = Val(tEArticulo.Tag)
    pcurQ_D = CCur(tEQ.Text)
    pdtmFecha_D = CDate(tEFecha.Text)
    pdblCosto_D = CCur(tECosto.Text)

    pdtmFecha_A = CDate("01/01/1900"): plngArticulo_A = 0: pcurQ_A = 0: pdblCosto_A = 0
    If Not CasoNuevo Then   'Está modificando !!
        'Saco los datos de la grilla como los Actuales = Antes  --
        pdtmFecha_A = CDate(vsConsulta.Cell(flexcpText, Row, 3))
        plngArticulo_A = vsConsulta.Cell(flexcpData, Row, 0)
        pcurQ_A = vsConsulta.Cell(flexcpText, Row, 4)
        pdblCosto_A = vsConsulta.Cell(flexcpText, Row, 5)
    End If
   
    
    cons = "Select * From CMAjusteExistencia Where AjEArticulo = 0"
    Set rsAj = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
   
    If CasoNuevo Or pcurQ_D Or ((plngArticulo_A = plngArticulo_D) And (pdtmFecha_A = pdtmFecha_D)) Then
          
        rsAj.AddNew
        rsAj!AjEFechaCambio = pdatCambio
        rsAj!AjETipo = mTipoCV
        rsAj!AjECodigo = mCodigo
        
        rsAj!AjEFecha = pdtmFecha_A
        
        If CasoNuevo Then
            rsAj!AjEArticulo = Val(tEArticulo.Tag)
        Else
            rsAj!AjEArticulo = plngArticulo_A
        End If
        
        rsAj!AjEQAntes = pcurQ_A
        rsAj!AjEQDespues = pcurQ_D
        rsAj!AjECostoAntes = pdblCosto_A
        rsAj!AjECostoDespues = pdblCosto_D
        If Trim(cComentario.Text) <> "" Then rsAj!AjeComentario = Trim(cComentario.Text)
        rsAj!AjEUsuario = paCodigoDeUsuario
        rsAj.Update
        
    Else
        'Ago Alta y Baja
        rsAj.AddNew
        rsAj!AjEFechaCambio = pdatCambio
        rsAj!AjETipo = mTipoCV
        rsAj!AjECodigo = mCodigo
        rsAj!AjEFecha = pdtmFecha_A
        rsAj!AjEArticulo = plngArticulo_A
        rsAj!AjEQAntes = pcurQ_A
        rsAj!AjEQDespues = 0
        rsAj!AjECostoAntes = pdblCosto_A
        rsAj!AjECostoDespues = 0
        If Trim(cComentario.Text) <> "" Then rsAj!AjeComentario = Trim(cComentario.Text)
        rsAj!AjEUsuario = paCodigoDeUsuario
        rsAj.Update
        
        rsAj.AddNew
        rsAj!AjEFechaCambio = pdatCambio
        rsAj!AjETipo = mTipoCV
        rsAj!AjECodigo = mCodigo

        rsAj!AjEFecha = pdtmFecha_D
        rsAj!AjEArticulo = plngArticulo_D
        rsAj!AjEQAntes = 0
        rsAj!AjEQDespues = pcurQ_D
        rsAj!AjECostoAntes = 0
        rsAj!AjECostoDespues = pdblCosto_D
        If Trim(cComentario.Text) <> "" Then rsAj!AjeComentario = Trim(cComentario.Text)
        rsAj!AjEUsuario = paCodigoDeUsuario
        rsAj.Update
    End If
    
    rsAj.Close
    Exit Function

errAjuste:
    clsGeneral.OcurrioError "Error al procesar el ajuste de existencia.", Err.Number & "- " & Err.Description
End Function

 
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
    sNuevo = False
    
    Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
    Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
    Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
    
    CargoCombo "Select GruCodigo as Codigo, GRuNombre as Nombre from Grupo Order by GruNombre", cGrupos
    
    cons = "SELECT RTrim(AjeComentario) From CMAjusteExistencia " & _
          " Where AjeComentario Is Not Null " & _
          " GROUP BY AjeComentario " & _
          " Having Count(*) > 2" & _
          " ORDER BY AjeComentario"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not rsAux.EOF
        cComentario.AddItem Trim(rsAux(0))
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Error al inicializar el formulario.", Trim(Err.Description)
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
   
    
    cons = "Select ArtID, ArtCodigo, ArtNombre, 'IDCompra' = Compra.ComCodigo, 'Fecha' = CM.ComFecha, CM.ComCodigo, ComCantidad, ComCosto, CM.ComTipo, ComTC " _
        & " From Articulo, CMCompra CM" _
                    & " Left Outer Join Compra On Compra.ComCodigo = CM.ComCodigo" _
            & " Where ComArticulo = ArtID "
    If IsDate(tFecha.Text) Then
        cons = cons & " And CM.ComFecha >= '" & Format(tFecha.Text & " 00:00:00", sqlFormatoFH) & "'" _
                           & " And CM.ComFecha <= '" & Format(tFHasta.Text & " 23:59:59", sqlFormatoFH) & "'"
    End If
    If Val(tArticulo.Tag) > 0 Then
        cons = cons & " And ComArticulo = " & CLng(tArticulo.Tag)
    Else
        If Trim(tNombres.Text) <> "" Then cons = cons & " And ArtNombre like '" & Trim(tNombres.Text) & "%'"
        If cGrupos.ListIndex <> -1 Then
            cons = cons & " And ComArticulo IN (Select AGrArticulo From ArticuloGrupo Where AGrGrupo = " & cGrupos.ItemData(cGrupos.ListIndex) & " )"
        End If
    End If
    cons = cons & " Order by ArtNombre, Fecha DESC"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If rsAux.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        rsAux.Close: Screen.MousePointer = 0: InicializoGrillas: Exit Sub
    End If
    Dim aIdAnterior As Long: aIdAnterior = 0
    Dim aValor As Long
    With vsConsulta
        .Rows = 1
        Do While Not rsAux.EOF
            'Data -->   0)Id_articulo     2)Id_Compra     3)ComTipo
                        
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = Format(rsAux!ArtCodigo, "(#,000,000)") & " " & Trim(rsAux!ArtNombre)
            aValor = rsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
            
            If rsAux!ComTipo = TipoCV.Compra Then .Cell(flexcpText, .Rows - 1, 2) = rsAux!IdCompra
            aValor = rsAux!ComCodigo: .Cell(flexcpData, .Rows - 1, 2) = aValor
            
            .Cell(flexcpText, .Rows - 1, 3) = Format(rsAux!Fecha, "dd/mm/yy")
            aValor = rsAux!ComTipo: .Cell(flexcpData, .Rows - 1, 3) = aValor
            
            .Cell(flexcpText, .Rows - 1, 4) = rsAux!ComCantidad
            .Cell(flexcpText, .Rows - 1, 5) = Format(rsAux!ComCosto, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 6) = Format(rsAux!ComCantidad * rsAux!ComCosto, FormatoMonedaP)
            If rsAux!ComTipo = TipoCV.Compra Then
                .Cell(flexcpText, .Rows - 1, 7) = Format(rsAux!ComTC, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 8) = Format(CCur(.Cell(flexcpText, .Rows - 1, 6)) / rsAux!ComTC, FormatoMonedaP)
            End If
            rsAux.MoveNext
        Loop
        rsAux.Close
        
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
    sNuevo = False
    tEArticulo.Enabled = Estado: tEQ.Enabled = Estado: tEFecha.Enabled = Estado: tECosto.Enabled = Estado: cComentario.Enabled = Estado
    
    If Estado Then
        tEArticulo.BackColor = Colores.Blanco: tEQ.BackColor = Colores.Blanco
        tEFecha.BackColor = Colores.Blanco: tECosto.BackColor = Colores.Blanco: cComentario.BackColor = Colores.Blanco
        tEArticulo.Text = vsConsulta.Cell(flexcpText, Row, 1): tEArticulo.Tag = vsConsulta.Cell(flexcpData, Row, 0)
        tEQ.Text = vsConsulta.Cell(flexcpText, Row, 4)
        tEFecha.Text = vsConsulta.Cell(flexcpText, Row, 3)
        tECosto.Text = vsConsulta.Cell(flexcpText, Row, 5)
        vsConsulta.Enabled = False: vsConsulta.BackColor = Colores.Inactivo
        Foco tEArticulo
    Else
        tEArticulo.BackColor = Colores.Inactivo
        tEQ.BackColor = Colores.Inactivo
        tEFecha.BackColor = Colores.Inactivo: tECosto.BackColor = Colores.Inactivo: cComentario.BackColor = Colores.Inactivo
        tEArticulo.Text = "": tEQ.Text = "": tEFecha.Text = "": tECosto.Text = "": cComentario.Text = ""
        vsConsulta.Enabled = True: vsConsulta.BackColor = Colores.Blanco
        vsConsulta.SetFocus
        
    End If
    
    vsConsulta.Tag = Row
    
End Sub

Private Sub EditoRenglonParaNuevo(Estado As Boolean, Optional Row As Long = 0)
    
    On Error Resume Next
    sNuevo = Estado
    tEArticulo.Enabled = Estado: tEQ.Enabled = Estado: tEFecha.Enabled = Estado: tECosto.Enabled = Estado: cComentario.Enabled = Estado
    
    If Estado Then
        tEArticulo.BackColor = Colores.Blanco: tEQ.BackColor = Colores.Blanco
        tEFecha.BackColor = Colores.Blanco: tECosto.BackColor = Colores.Blanco: cComentario.BackColor = Colores.Blanco
        tEArticulo.Text = "": tEQ.Text = "": tEFecha.Text = "": tECosto.Text = "": cComentario.Text = ""
        
        vsConsulta.Enabled = False: vsConsulta.BackColor = Colores.Inactivo
        Foco tEArticulo
    Else
        tEArticulo.BackColor = Colores.Inactivo
        tEQ.BackColor = Colores.Inactivo
        tEFecha.BackColor = Colores.Inactivo: tECosto.BackColor = Colores.Inactivo: cComentario.BackColor = Colores.Inactivo
        tEArticulo.Text = "": tEQ.Text = "": tEFecha.Text = "": tECosto.Text = "": cComentario.Text = ""
        vsConsulta.Enabled = True: vsConsulta.BackColor = Colores.Blanco
        vsConsulta.SetFocus
    End If
    
    'vsConsulta.Tag = Row
    
End Sub


Private Sub tNombres_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco bConsultar
End Sub

Private Sub vsConsulta_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And vsConsulta.Rows > 1 Then
        sNuevo = False
        EditoRenglon True, vsConsulta.Row
    End If
    
End Sub

Private Function AgregoComentario()
    
    If cComentario.ListIndex = -1 And Trim(cComentario.Text) <> "" Then
        cComentario.AddItem Trim(cComentario.Text)
        
        cComentario.ListIndex = cComentario.ListCount - 1
    End If

End Function

