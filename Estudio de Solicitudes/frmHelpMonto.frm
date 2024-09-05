VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form frmHelpMonto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculos Financiando "
   ClientHeight    =   4245
   ClientLeft      =   3315
   ClientTop       =   3345
   ClientWidth     =   6120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHelpMonto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6120
   Begin VB.CommandButton bCancel 
      Caption         =   "Cancelar"
      Height          =   325
      Left            =   4980
      TabIndex        =   2
      Top             =   3850
      Width           =   1095
   End
   Begin VB.CommandButton bOK 
      Caption         =   "Aceptar"
      Height          =   325
      Left            =   3780
      TabIndex        =   1
      Top             =   3850
      Width           =   1095
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   2778
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
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   4
      GridLinesFixed  =   2
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
End
Attribute VB_Name = "frmHelpMonto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public prmIDSolicitud As Long
Public prmTotalFinanciado As Currency
Public prmSugerirPlanes As String

Private mSQL As String
Private rsQry As rdoResultset

Private Sub bCancel_Click()
prmSugerirPlanes = ""
Unload Me
End Sub

Private Sub bOK_Click()

On Error GoTo errUnload
Dim idX As Integer
    
    With vsLista
        For idX = .FixedRows To .Rows - .FixedRows
            If .Cell(flexcpChecked, idX, 0) = flexChecked Then
                prmSugerirPlanes = prmSugerirPlanes & IIf(prmSugerirPlanes <> "", "; ", "")
                If Trim(.Cell(flexcpText, idX, 2)) <> "" Then
                    prmSugerirPlanes = prmSugerirPlanes & "E:" & Trim(.Cell(flexcpText, idX, 2)) & "+" & Trim(.Cell(flexcpText, idX, 3))
                Else
                      prmSugerirPlanes = prmSugerirPlanes & Trim(.Cell(flexcpText, idX, 3))
                End If
            End If
        Next
    End With
    If prmSugerirPlanes = "" Then prmSugerirPlanes = prmTotalFinanciado
    Unload Me
    
errUnload:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Call bCancel_Click
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    prmSugerirPlanes = ""

    zfn_InicializoControles
    fnc_CalculoPrecios
    Me.Caption = "Cálculos Financiando " & Format(prmTotalFinanciado, "#,##0")
    
    Screen.MousePointer = 0
    
End Sub

Private Function fnc_CalculoPrecios()
On Error GoTo errSQL
'Dim bOK  As Boolean

    Screen.MousePointer = 11
'    Cons = "Select  TCuCodigo, Count(Distinct(CoeCoeficiente))" & _
'            " From PrecioVigente, Coeficiente, TipoCuota " & _
'            " Where PVIArticulo IN (Select RSoArticulo from RenglonSolicitud Where RSoSolicitud = " & prmIDSolicitud & ")" & _
'            " And PViMoneda = " & paMonedaPesos & _
'            " And PViTipoCuota = " & paTipoCuotaContado & _
'            " And PViHabilitado = 1 And TCuDeshabilitado Is Null And TCuEspecial = 0 And TCuLlevaConforme > 0 " & _
'            " And PViPlan = CoePlan And TCuCodigo = CoeTipoCuota  And CoeMoneda = PViMoneda " & _
'            " And TCuVencimientoE = 0 " & _
'            " Group by TCuCodigo" & _
'            " Having Count(Distinct(CoeCoeficiente)) > 1"
'
'    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'    bOK = RsAux.EOF
'    RsAux.Close
    
'    If bOK Then          'Si no hay más de un Plan       -------------------------------
    
'    mSQL = " Select  CoeCoeficiente, TCuCantidad as Cuotas,  TCuAbreviacion as Financiacion,Sum(PViPrecio) as Contado," & _
                    " Sum(PViPrecio)-(" & prmTotalFinanciado & "/CoeCoeficiente) as EInicial, " & _
                    " ((Sum(PViPrecio)-(Sum(PViPrecio)-(" & prmTotalFinanciado & "/CoeCoeficiente)))*CoeCoeficiente) / TCuCantidad as VCuota " & _
                " From PrecioVigente, Coeficiente, TipoCuota " & _
                " Where PVIArticulo IN (Select RSoArticulo from RenglonSolicitud Where RSoSolicitud = " & prmIDSolicitud & ")" & _
                " And PViMoneda = " & paMonedaPesos & _
                " And PViTipoCuota = " & paTipoCuotaContado & _
                " And PViHabilitado = 1 And TCuDeshabilitado Is Null And TCuEspecial = 0 And TCuLlevaConforme > 0 " & _
                " And PViPlan = CoePlan And TCuCodigo = CoeTipoCuota  And CoeMoneda = PViMoneda " & _
                " And TCuVencimientoE = 0 " & _
                " Group by TCuOrden, CoeCoeficiente, TCuCantidad,TCuAbreviacion " & _
                " Order BY TCuOrden"


    '(planes diferentes)
    Dim mValor As Currency
    mSQL = "EXEC prg_PlanesSolicitudHasta " & prmIDSolicitud & ", " & prmTotalFinanciado
    Set rsQry = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    
    If Not rsQry.EOF Then prmTotalFinanciado = rsQry("Hasta")
    Do While Not rsQry.EOF
        With vsLista
            .AddItem ""
            .Cell(flexcpChecked, .Rows - 1, 0) = flexUnchecked
            
            .Cell(flexcpText, .Rows - 1, 1) = Trim(rsQry!Financiacion)
            mValor = rsQry!Cuotas: .Cell(flexcpData, .Rows - 1, 1) = mValor
            
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsQry!EInicial, "#,##0")
            
            .Cell(flexcpText, .Rows - 1, 3) = rsQry!Cuotas & "x " & Format(rsQry!VCuota, "#,##0")
            mValor = rsQry!VCuota: .Cell(flexcpData, .Rows - 1, 3) = mValor
            
            .Cell(flexcpText, .Rows - 1, 4) = Format(rsQry!EInicial + (rsQry!Cuotas * Format(rsQry!VCuota, "#,##0")), "#,##0")
            
'            .Cell(flexcpText, .Rows - 1, 5) = Format(rsQry!Contado, "#,##0")
'            .Cell(flexcpText, .Rows - 1, 6) = rsQry!CoeCoeficiente
            
            
            .Cell(flexcpBackColor, .Rows - 1, 2, , 3) = RGB(70, 130, 180) '&H808000   '&HFFC0C0 '&HFFC8B5
            .Cell(flexcpForeColor, .Rows - 1, 2, , 3) = vbWhite
        End With
        rsQry.MoveNext
    Loop
    rsQry.Close
    
    
'    With vsLista    'SI EL valor de la cuota no da multiplo de 10   --> XX/10 --Redondeo y multiuplico *10
'        Dim mCuota As Currency, mTotalCtas As Currency, mEntrega As Currency
'        Dim idX As Integer
'
'        For idX = .FixedRows To .Rows - .FixedRows
'            mCuota = .Cell(flexcpData, idX, 3)
'            If (mCuota Mod 10) <> 0 Then
'                mCuota = Round(mCuota / 10, 0) * 10
'                mTotalCtas = mCuota * .Cell(flexcpData, idX, 1)
'                mEntrega = Format(.Cell(flexcpValue, idX, 5) - (mTotalCtas / .Cell(flexcpValue, idX, 6)), "#,##0")
'
'                If mTotalCtas > prmTotalFinanciado Then prmTotalFinanciado = mTotalCtas
'
'                .Cell(flexcpText, idX, 2) = Format(mEntrega, "#,##0")
'                .Cell(flexcpText, idX, 3) = .Cell(flexcpData, idX, 1) & "x " & Format(mCuota, "#,##0")
'                .Cell(flexcpText, idX, 4) = Format(mEntrega + mTotalCtas, "#,##0")
'            End If
'
'        Next
'    End With
    
'    Else
'        With vsLista
'            .AddItem ""
'            .Cell(flexcpChecked, .Rows - 1, 0) = flexNoCheckbox
'            .Cell(flexcpText, .Rows - 1, 1) = " ... hay varios planes ..."
'        End With
'    End If
    
    '2) Calculo los precios Para las cuotas sin entrega     -------------------------------------------------------
'    mSQL = "Select  TCuCantidad as Cuotas,  TCuAbreviacion as Financiacion, Sum(PViPrecio / TCuCantidad) As VCuota" & _
'                " From PrecioVigente, TipoCuota" & _
'                " Where PVIArticulo IN (Select RSoArticulo from RenglonSolicitud Where RSoSolicitud = " & prmIDSolicitud & ")" & _
'                " And PViTipoCuota = TCuCodigo " & _
'                " And TCuDeshabilitado Is Null And TCuVencimientoE Is Null  And TCuVencimientoC = 0 " & _
'                " And PViMoneda = " & paMonedaPesos & _
'                " And PViTipoCuota <> " & paTipoCuotaContado & _
'                " And PViHabilitado = 1 " & _
'                " Group by TCuOrden, TCuCantidad,TCuAbreviacion " & _
'                "Order BY TCuOrden "
'
'    Set rsQry = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
'    Do While Not rsQry.EOF
'        If ((rsQry!Cuotas - 1) * rsQry!VCuota) <= prmTotalFinanciado Then
'            With vsLista
'                .AddItem ""
'                .Cell(flexcpChecked, .Rows - 1, 0) = flexUnchecked
'
'                .Cell(flexcpText, .Rows - 1, 1) = Trim(rsQry!Financiacion)
'                mValor = rsQry!Cuotas: .Cell(flexcpData, .Rows - 1, 1) = mValor
'
'                .Cell(flexcpText, .Rows - 1, 3) = rsQry!Cuotas & "x " & Format(rsQry!VCuota, "#,##0")
'                mValor = rsQry!VCuota: .Cell(flexcpData, .Rows - 1, 3) = mValor
'
'                .Cell(flexcpText, .Rows - 1, 4) = Format(rsQry!VCuota * rsQry!Cuotas, "#,##0")
'
'                .Cell(flexcpBackColor, .Rows - 1, 3, , 3) = RGB(46, 139, 87) '&H808000   '&HFFC0C0 '&HFFC8B5
'                .Cell(flexcpForeColor, .Rows - 1, 3, , 3) = vbWhite
'            End With
'        End If
'        rsQry.MoveNext
'    Loop
'    rsQry.Close
    Screen.MousePointer = 0
    Exit Function
    
errSQL:
    clsGeneral.OcurrioError "Error al ejecutar la consulta para sugerir financiaciones.", Err.Description
End Function

Private Function zfn_InicializoControles()

    With vsLista
        .Rows = 1: .Cols = 1
        .FormatString = "^Sugerir|<Cuotas|>Entrega de|>Cuotas de|>Precio Final" '|>Contado|<Coeficiente"
        .ColWidth(1) = 800: .ColWidth(2) = 1000: .ColWidth(3) = 1300: .ColWidth(4) = 1100 ': .ColWidth(5) = 1000
        '.ColHidden(6) = True
        
        .WordWrap = False
        '.MergeCells = flexMergeNever '= flexMergeSpill
        .MergeCells = flexMergeSpill
        .ExtendLastCol = True
        .Editable = True
        
        .BackColorBkg = .BackColor
        .SheetBorder = .BackColor
        .GridLines = flexGridFlat
    End With
    
End Function

Private Sub Form_Resize()
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    vsLista.Move 20, 20, Me.ScaleWidth - (20 * 2), Me.ScaleHeight - (150 + bOK.Height)
    
End Sub

Private Sub vsLista_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col > 0 Then Cancel = True: Exit Sub
End Sub

Private Sub vsLista_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bOK.SetFocus
End Sub
