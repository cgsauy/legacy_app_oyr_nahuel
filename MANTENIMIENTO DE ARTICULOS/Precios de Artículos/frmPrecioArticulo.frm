VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form frmPrecioArticulo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Precios de Artículos"
   ClientHeight    =   5505
   ClientLeft      =   2895
   ClientTop       =   2235
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrecioArticulo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6990
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir. [Ctrl+X]"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   500
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "refrescar"
            Object.ToolTipText     =   "Refrescar"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.OptionButton opVigencia 
      Caption         =   "A &Regir"
      Height          =   195
      Index           =   1
      Left            =   5460
      TabIndex        =   15
      Top             =   1800
      Width           =   1335
   End
   Begin VB.OptionButton opVigencia 
      Caption         =   "Vigen&tes"
      Height          =   195
      Index           =   0
      Left            =   4440
      TabIndex        =   14
      Top             =   1800
      Width           =   975
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsPrecio 
      Height          =   3075
      Left            =   120
      TabIndex        =   13
      Top             =   2100
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5424
      _ConvInfo       =   1
      Appearance      =   1
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
      SelectionMode   =   0
      GridLines       =   1
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
   Begin VB.TextBox tArticulo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   780
      TabIndex        =   1
      Top             =   600
      Width           =   4995
   End
   Begin AACombo99.AACombo cMoneda 
      Height          =   315
      Left            =   780
      TabIndex        =   7
      Top             =   1320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      BackColor       =   12648447
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
   Begin VB.TextBox tPlan 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   780
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox tValor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   10
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox tVigencia 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   4320
      MaxLength       =   16
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.CheckBox cHabilitado 
      Alignment       =   1  'Right Justify
      Caption         =   "&En Uso:"
      Height          =   255
      Left            =   6000
      TabIndex        =   11
      Top             =   1320
      Width           =   850
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   5250
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8625
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3625
            MinWidth        =   2
            Text            =   "Ctr+H - Historia de Precios "
            TextSave        =   "Ctr+H - Historia de Precios "
         EndProperty
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
   End
   Begin AACombo99.AACombo cCuota 
      Height          =   315
      Left            =   1620
      TabIndex        =   8
      Top             =   1320
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      BackColor       =   12648447
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7140
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrecioArticulo.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrecioArticulo.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrecioArticulo.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrecioArticulo.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrecioArticulo.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrecioArticulo.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrecioArticulo.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrecioArticulo.frx":0DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrecioArticulo.frx":10E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrecioArticulo.frx":13FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Lista de Precios"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   1740
      Width           =   6735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Artículo:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "&Cuotas:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "V&alor Cuota:"
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "&Plan:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "&Vigencia:"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuModificar 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuLinea2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRefrescar 
         Caption         =   "&Refrescar"
         Shortcut        =   {F5}
      End
      Begin VB.Menu MnuHistoria 
         Caption         =   "&Historia de Precios"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "&Del formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmPrecioArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Cambios
'25-10-2001
'    En distribuir cuotas cambiamos consulta para que no cargue los coeficientes = 1 y las Tipo cuotas que no estén habilitadas.

Private Const cteColorPV = &HC0C0FF
Private sNuevo As Boolean, sModificar As Boolean
Private sVigente As Boolean             'Indica si hay precios vigentes ingresados (para Hab/Des botones)

Private aFecha As Date        'Para controlar modificaciones multiusuario

Private Sub DistribuirCuotas(Plan As Integer, Moneda As Integer, Contado As Currency)
Dim bCargueElContado As Boolean     'Para saber si el contado fue ingresado (por si no pertenece al plan)
Dim lCont As Long

    bCargueElContado = False
    'Saco los datos de los coeficientes con = moneda a la del precio y sin entrega (TCuVencimientoE = NULL)
    Cons = "Select * from Coeficiente, TipoCuota" _
        & " Where CoePlan =  " & Plan _
        & " And CoeTipoCuota = TCuCodigo" _
        & " And CoeMoneda = " & Moneda _
        & " And TCuVencimientoE Is NULL And TCuDeshabilitado Is Null " _
        & " And CoeCoeficiente <> 1 Order by TCuOrden"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        
        If RsAux!TCuCodigo = paTipoCuotaContado Then bCargueElContado = True
        
        With vsPrecio
            'Verifico si hay ingresado un precio para la misma moneda---
            lCont = 1
            Do While lCont < .Rows - 1
                If CCur(.Cell(flexcpData, lCont, 0)) = RsAux!CoeTipoCuota And _
                    CCur(.Cell(flexcpData, lCont, 1)) = Moneda Then
                    .RemoveItem lCont
                Else
                    lCont = lCont + 1
                End If
            Loop
            '----------------------------------------------------------------------
        
            .AddItem Trim(RsAux!TCuAbreviacion)
            'Agrego.
            lCont = RsAux!TCuCodigo: .Cell(flexcpData, .Rows - 1, 0) = lCont
            lCont = RsAux!CoeMoneda: .Cell(flexcpData, .Rows - 1, 1) = lCont
            lCont = RsAux!TCuCantidad: .Cell(flexcpData, .Rows - 1, 2) = lCont
            
            .Cell(flexcpText, .Rows - 1, 1) = Trim(cMoneda.Text)
            .Cell(flexcpText, .Rows - 1, 3) = Format((Contado * RsAux!CoeCoeficiente) / RsAux!TCuCantidad, "#,##0") & ".00"
            .Cell(flexcpForeColor, .Rows - 1, 3) = ColorAjuste(.Cell(flexcpValue, .Rows - 1, 3))
            .Cell(flexcpText, .Rows - 1, 2) = Format(CCur(.Cell(flexcpText, .Rows - 1, 3)) * RsAux!TCuCantidad, "#,##0") & ".00"
            .Cell(flexcpText, .Rows - 1, 4) = Format(CDate(tVigencia.Text), "dd/mm/yy hh:nn")
            If cHabilitado.Value = 1 Then
                .Cell(flexcpChecked, .Rows - 1, 5) = flexChecked
            Else
                .Cell(flexcpChecked, .Rows - 1, 5) = flexUnchecked
            End If
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If Not bCargueElContado Then
        'Es un Contado que no pertenece al plan (la excepcion a la regla)
        AgregoCuota cCuota.ItemData(cCuota.ListIndex), Moneda
    End If
    
    'Ahora veo si en los precios anteriores tengo algún tipo de cuota habilitado que no este incluido
    ' en este nuevo ingreso.
    ControlPrecioNuevoContraVigentes
    
    
End Sub

Private Sub AgregoCuota(Cuota As Integer, Moneda As Integer)

Dim aCantidad As Integer

    Cons = "Select * From TipoCuota Where TCuCodigo = " & Cuota
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    aCantidad = RsAux!TCuCantidad
    RsAux.Close
    
    With vsPrecio
        .AddItem Trim(cCuota.Text)
        .Cell(flexcpData, .Rows - 1, 0) = Cuota
        .Cell(flexcpData, .Rows - 1, 1) = Moneda
        .Cell(flexcpData, .Rows - 1, 2) = aCantidad
        
        .Cell(flexcpText, .Rows - 1, 1) = Trim(cMoneda.Text)
        .Cell(flexcpText, .Rows - 1, 3) = Format(tValor.Text, "#,##0.00")       'Valor de la cuota
        .Cell(flexcpForeColor, .Rows - 1, 3) = ColorAjuste(.Cell(flexcpValue, .Rows - 1, 3))
        .Cell(flexcpText, .Rows - 1, 2) = Format(CCur(.Cell(flexcpText, .Rows - 1, 3)) * aCantidad, "#,##0.00")
        .Cell(flexcpText, .Rows - 1, 4) = Format(tVigencia.Text, "dd/mm/yy hh:mm")
        If cHabilitado.Value = 1 Then
            .Cell(flexcpChecked, .Rows - 1, 5) = flexChecked
        Else
            .Cell(flexcpChecked, .Rows - 1, 5) = flexUnchecked
        End If
    End With
    
End Sub

Private Sub cCuota_GotFocus()
    With cCuota
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Status.Panels(1).Text = "Seleccione el tipo de cuota de financiación."
End Sub

Private Sub cCuota_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tValor
End Sub
Private Sub cCuota_LostFocus()
    cCuota.SelLength = 0
    Status.Panels(1).Text = ""
End Sub
Private Sub cHabilitado_GotFocus()

    Status.Panels(1).Text = "Indique si el precio está en uso."
    
End Sub

Private Sub cHabilitado_KeyPress(KeyAscii As Integer)
Dim lCont As Long
    If KeyAscii = vbKeyReturn Then
        If Not sNuevo And Not sModificar Then Exit Sub
        On Error GoTo errAgregar
             
        'Valido que se hayan seleccionado los datos necesarios-----------------------------------------------
        If Val(tPlan.Tag) = 0 Then
            MsgBox "Seleccione un plan para realizar la distribución.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
             
        If Not IsNumeric(tValor.Text) Or Trim(tValor.Text) = "" Or cMoneda.ListIndex = -1 Or Not IsDate(tVigencia.Text) Then
            MsgBox "Los datos ingresados no son correctos.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
        '--------------------------------------------------------------------------------------------------------------
        
        'Agrego los valores a la lista
        If sNuevo Then  'TRABAJO CON LOS PRECIOS A REGIR-----------------------------------------------------------------
            With vsPrecio
                For lCont = 1 To .Rows - 1
                    If Val(.Cell(flexcpData, lCont, 0)) = Val(cCuota.ItemData(cCuota.ListIndex)) And _
                        Val(.Cell(flexcpData, lCont, 0)) = cMoneda.ItemData(cMoneda.ListIndex) Then
                    MsgBox "Existe un precio para la moneda y tipo de cuota seleccionados, para modificarlo edítelo.", vbExclamation, "ATENCIÓN"
                    Exit Sub
                    End If
                Next
            End With
            '----------------------------------------------------------------------
            If Format(tVigencia.Text, "yyyy/mm/dd hh:mm") <= Format(Now, "yyyy/mm/dd hh:mm") Then
                MsgBox "La fecha de vigencia ingresada no debe ser menor a la actual.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
        
            If cCuota.ItemData(cCuota.ListIndex) = paTipoCuotaContado Then
                DistribuirCuotas tPlan.Tag, cMoneda.ItemData(cMoneda.ListIndex), CCur(tValor.Text)
            Else
                AgregoCuota cCuota.ItemData(cCuota.ListIndex), cMoneda.ItemData(cMoneda.ListIndex)
            End If
                        
            cMoneda.Text = ""
            cCuota.Text = ""
            tValor.Text = ""
            cHabilitado.Value = 0
            Foco cMoneda
            
        End If
            
        'Accion Modificar   'TRABAJO CON LOS PRECIOS VIGENTES----------------------------------------------------------------
        If sModificar Then
            
            With vsPrecio
                'Valor de la Cuota
                .Cell(flexcpText, .Row, 3) = Format(CCur(tValor.Text), "#,##0") & ".00"
                .Cell(flexcpForeColor, .Row, 3) = ColorAjuste(.Cell(flexcpValue, .Row, 3))
                .Cell(flexcpText, .Row, 2) = Format(CCur(.Cell(flexcpText, .Row, 3)) * Val(.Cell(flexcpData, .Row, 2)), "#,##0") & ".00"
                If cHabilitado.Value = 0 Then
                    .Cell(flexcpChecked, .Row, 5) = flexUnchecked
                Else
                    .Cell(flexcpChecked, .Row, 5) = flexChecked
                End If
            End With
            LimpioCampos
            tValor.Enabled = False: tValor.BackColor = vbButtonFace
            cHabilitado.Enabled = False
            vsPrecio.SetFocus
        End If
    End If
    Exit Sub

errAgregar:
    clsGeneral.OcurrioError "Ocurrió un error al agregar el tipo de cuota. Verifique que no esté ingresado.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub cMoneda_GotFocus()
    cMoneda.SelStart = 0
    cMoneda.SelLength = Len(cMoneda.Text)
    Status.Panels(1).Text = "Seleccione la moneda para el ingreso de precios."
End Sub
Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cCuota
End Sub
Private Sub cMoneda_LostFocus()
    cMoneda.SelLength = 0
    Status.Panels(1).Text = ""
End Sub

Private Sub Form_Activate()

    Screen.MousePointer = 11
    Screen.MousePointer = 0
    DoEvents

End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    
    ObtengoSeteoForm Me, 500, 500
    sNuevo = False: sModificar = False
    Me.Height = 6165
    With vsPrecio
        .Editable = False
        .Rows = 1: .Cols = 1: .ExtendLastCol = True
        .FormatString = "<Tipo de Cuota|<Moneda|Precio|Valor Cuota|>Vigencia al|En Uso"
        .ColWidth(2) = 1350
        .ColWidth(3) = 1350
        .ColWidth(4) = 1350
        .ColDataType(.Cols - 1) = flexDTBoolean
        .ColDataType(4) = flexDTDate
        '.ColFormat(4) = "dd/mm/yyyy hh:nn:ss"
    End With
    DeshabilitoIngreso
    
    'Cargo las MONEDAS
    Cons = "Select MonCodigo, MonSigno From Moneda Order by MonSigno"
    CargoCombo Cons, cMoneda, ""
    
    'Cargo los Tipos de Cuotas
    Cons = "Select TCuCodigo, TCuAbreviacion from TipoCuota" _
            & " Where TCuVencimientoE = NULL"
    CargoCombo Cons, cCuota, ""
    
    If lArtID = 0 Then
        Botones False, False, False, False, False, Toolbar1, Me
    Else
        tArticulo.Tag = lArtID
        CargoArticuloPorID
    End If
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al ingresar al formulario."
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If sNuevo Or sModificar Then
        If MsgBox("Ud. realizó modificaciones en la ficha y no ha grabado." & Chr(13) _
            & "Desea almacenar la información ingresada.", vbYesNo + vbExclamation, "ATENCIÓN") = vbYes Then
            
            AccionGrabar
            If sNuevo Or sModificar Then
                Cancel = True
                Exit Sub
            End If
            
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    
    GuardoSeteoForm Me
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
    Exit Sub
    
End Sub

Private Sub Label17_Click()
    Foco tVigencia
End Sub

Private Sub Label2_Click()
On Error Resume Next
    With tArticulo
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label25_Click()
On Error Resume Next
    With tPlan
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label26_Click()
    On Error Resume Next
    With tValor
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label27_Click()
    On Error Resume Next
    With cMoneda
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuEliminar_Click()
    AccionEliminar
End Sub

Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuHistoria_Click()
On Error GoTo ErrMH
    
    Screen.MousePointer = 0
    'CONSULTA DE HISTORIAS DE PRECIOS
    Cons = "Select  'Plan' = PlaNombre,   'Tipo de Cuota' = TCuAbreviacion, Moneda = MonSigno, Precio = HPrPrecio, Vigencia = HPrVigencia" _
           & " From HistoriaPrecio, Moneda, TipoCuota, TipoPlan" _
           & " Where HPrMoneda = MonCodigo" _
           & " And HPrTipoCuota = TCuCodigo" _
           & " And HPrArticulo = " & Val(tArticulo.Tag) & " And HPrPlan = PlaCodigo " _
           & "Order by HPrVigencia DESC"
           
    Dim objAyuda As New clsListadeAyuda
    objAyuda.ActivarAyuda cBase, Cons, 7000, 0
    Set objAyuda = Nothing
    Exit Sub
    
ErrMH:
    clsGeneral.OcurrioError "Ocurrio un error inesperado al presentar la lista de ayuda.", Err.Description
End Sub

Private Sub MnuModificar_Click()
    AccionModificar
End Sub

Private Sub MnuNuevo_Click()
    AccionNuevo
End Sub

Private Sub MnuRefrescar_Click()
    AccionRefrescar
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Sub AccionNuevo()
On Error GoTo ErrAN
    Screen.MousePointer = 11
    
    'Prendo Señal que es uno nuevo.
    sNuevo = True
    
    'Habilito y Desabilito Botones.
    Call Botones(False, False, False, True, True, Toolbar1, Me)
    CargoDatosArticuloARegir
    HabilitoIngreso
    opVigencia(0).Enabled = False
    opVigencia(1).Enabled = False
    Foco tPlan
    vsPrecio.Editable = True
    Screen.MousePointer = 0
    Exit Sub
    
ErrAN:
    clsGeneral.OcurrioError "Ocurrio un error inesperado.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub LimpioCampos()
    Status.Panels(1).Text = ""
    cMoneda.Text = ""
    cCuota.Text = ""
    tValor.Text = ""
    tVigencia.Text = ""
    cHabilitado.Value = 0
End Sub

Sub AccionModificar()

    'Prendo Señal que es modificación.
    sModificar = True
    
    'Habilito y Desabilito Botones.
    Call Botones(False, False, False, True, True, Toolbar1, Me)
    HabilitoIngreso
    Screen.MousePointer = 11
    If opVigencia(1).Value Then
        MsgBox "Los precios que se pueden modificar son solo los vigentes.", vbInformation, "ATENCIÓN"
    End If
    CargoDatosArticuloVigente
    opVigencia(0).Enabled = False
    opVigencia(1).Enabled = False
    Status.Panels(1).Text = "Seleccione un precio y de enter para editar."
    
    With vsPrecio
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, 2) = vbButtonFace
        .Cell(flexcpBackColor, 1, 4, .Rows - 1, 4) = vbButtonFace
        .Editable = True
        .SetFocus
    End With
    Screen.MousePointer = 0

End Sub

Sub AccionGrabar()
Dim bSuceso As Boolean
Dim iCont As Integer
Dim lUID As Long, sDefensa As String, sDescSuceso As String
Dim objSuceso As clsSuceso
    If Not ValidoCampos Then Exit Sub
        
    If MsgBox("Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then
        Screen.MousePointer = 11
        
        sDescSuceso = "Modificación de Precios Vigente."
        
        If sNuevo Then      'NUEVO ARTICULO
            
            On Error GoTo ErrBT
            'Veo si hay suceso.
            For iCont = 1 To vsPrecio.Rows - 1
            
                If vsPrecio.Cell(flexcpBackColor, iCont) = cteColorPV Then
                
                    Cons = "Select * from HistoriaPrecio" _
                        & " Where HPrArticulo = " & Val(tArticulo.Tag) _
                        & " And HPrTipoCuota = " & Val(vsPrecio.Cell(flexcpData, iCont, 0)) _
                        & " And HPrMoneda = " & Val(vsPrecio.Cell(flexcpData, iCont, 1)) _
                        & " And HPrVigencia = '" & Format(vsPrecio.Cell(flexcpText, iCont, 4), "mm/dd/yyyy hh:nn:00") & "'" _
            
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    
                    If Not RsAux.EOF Then
                        If RsAux!HPrPrecio <> vsPrecio.Cell(flexcpText, iCont, 2) Then
 '                           sDescSuceso = sDescSuceso & vbCrLf & "Valor Antes: " & Trim(vsPrecio.Cell(flexcpText, iCont, 0)) & " " & Trim(vsPrecio.Cell(flexcpText, iCont, 1)) & " " & Trim(RsAux!HPrPrecio) & _
                                vbCrLf & "Valor Nuevo: " & Trim(vsPrecio.Cell(flexcpText, iCont, 0)) & " " & Trim(vsPrecio.Cell(flexcpText, iCont, 1)) & " " & Trim(vsPrecio.Cell(flexcpText, iCont, 1))
                            RsAux.Close: bSuceso = True: Exit For
                        End If
                    End If
                    RsAux.Close
                End If
            Next
            
            lUID = 0
            If bSuceso Then
                Set objSuceso = New clsSuceso
                objSuceso.TipoSuceso = TipoSuceso.Varios
                objSuceso.ActivoFormulario miConexion.UsuarioLogueado(True), "Cambio de Precio de Artículo", cBase
                lUID = objSuceso.Usuario
                sDefensa = objSuceso.Defensa
                Me.Refresh
                Set objSuceso = Nothing
                If lUID = 0 Then Screen.MousePointer = 0: Exit Sub
            End If
            
            cBase.BeginTrans    'COMIENZO LA TRANSACCION--------------------------
            On Error GoTo ErrET
            
            'Cargo los datos en HistoriaPrecio
            CargoCamposBDNuevo
            
            If bSuceso Then
                clsGeneral.RegistroSuceso cBase, Now, TipoSuceso.Varios, paCodigoDeTerminal, lUID, 0, Val(tArticulo.Tag), sDescSuceso, sDefensa
            End If
            
            cBase.CommitTrans   'FIN DE LA TRANSACCION-------------------------------
            sNuevo = False
            'Veo Si tiene precios vigentes para habilitar boton--------------
            Cons = "Select * from PrecioVigente Where PViArticulo = " & tArticulo.Tag
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                sVigente = True
            Else
                sVigente = False
            End If
            RsAux.Close '--------------------------------------------------------
            Botones True, sVigente, True, False, False, Toolbar1, Me
            
        Else                      'MODIFICACION DE LOS DATOS
            
            For iCont = 1 To vsPrecio.Rows - 1
                
                Cons = "Select * from HistoriaPrecio" _
                    & " Where HPrArticulo = " & Val(tArticulo.Tag) _
                    & " And HPrTipoCuota = " & Val(vsPrecio.Cell(flexcpData, iCont, 0)) _
                    & " And HPrMoneda = " & Val(vsPrecio.Cell(flexcpData, iCont, 1)) _
                    & " And HPrVigencia = '" & Format(vsPrecio.Cell(flexcpText, iCont, 4), "mm/dd/yyyy hh:nn:00") & "'" _
        
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                
                If Not RsAux.EOF Then
                    If RsAux!HPrPrecio <> CCur(vsPrecio.Cell(flexcpText, iCont, 2)) Then
'                        sDescSuceso = sDescSuceso & vbCrLf & "Valor Antes: " & Trim(vsPrecio.Cell(flexcpText, iCont, 0)) & " " & Trim(vsPrecio.Cell(flexcpText, iCont, 1)) & " " & Trim(RsAux!HPrPrecio) & _
                            vbCrLf & "Valor Nuevo: " & Trim(vsPrecio.Cell(flexcpText, iCont, 0)) & " " & Trim(vsPrecio.Cell(flexcpText, iCont, 1)) & " " & Trim(vsPrecio.Cell(flexcpText, iCont, 1))
                        RsAux.Close: bSuceso = True: Exit For
                    End If
                End If
                RsAux.Close
                
            Next
            
            lUID = 0
            If bSuceso Then
                Set objSuceso = New clsSuceso
                objSuceso.TipoSuceso = TipoSuceso.Varios
                objSuceso.ActivoFormulario miConexion.UsuarioLogueado(True), "Cambio de Precio de Artículo", cBase
                lUID = objSuceso.Usuario
                sDefensa = objSuceso.Defensa
                Me.Refresh
                Set objSuceso = Nothing
                If lUID = 0 Then Screen.MousePointer = 0: Exit Sub
            End If
            
            On Error GoTo ErrBT
            cBase.BeginTrans    'COMIENZO LA TRANSACCION--------------------------
            On Error GoTo ErrET
            
            'Cargo los datos en HistoriaPrecio
            CargoCamposBDModificar
            If bSuceso Then
                clsGeneral.RegistroSuceso cBase, Now, TipoSuceso.Varios, paCodigoDeTerminal, lUID, 0, Val(tArticulo.Tag), sDescSuceso, sDefensa
            End If
            cBase.CommitTrans   'FIN DE LA TRANSACCION-------------------------------
            
            sModificar = False
            Botones True, True, True, False, False, Toolbar1, Me
        End If
        
        DeshabilitoIngreso
        LimpioCampos
        Screen.MousePointer = 0
        
    End If
    Exit Sub
    
ErrBT:
    clsGeneral.OcurrioError "No se pudo iniciar la transacción.", Err.Description
    Screen.MousePointer = vbDefault
    Exit Sub
ErrET:
    Resume ErrRoll
ErrRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "No se pudo almacenar la información, reintente."
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Sub AccionEliminar()

    Screen.MousePointer = 11
    'Veo Si tiene precios A REGIR para eliminar
    Cons = "Select * from PrecioARegir Where PReArticulo = " & Val(tArticulo.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        Screen.MousePointer = 0
        MsgBox "No hay precios A Regir para eliminar.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    RsAux.Close
    
    Screen.MousePointer = 0
    CargoDatosArticuloARegir
    If MsgBox("Confrima eliminar la lista de precios a regir.", vbQuestion + vbYesNo, "ELIMINAR") = vbYes Then
        Screen.MousePointer = 11
        Cons = "Select * From Articulo Where ArtID = " & Val(tArticulo.Tag)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If aFecha <> RsAux!ArtModificado Then
            RsAux.Close
            Screen.MousePointer = 0
            MsgBox "La ficha ha sido modificada por otro usuario. Verifique los datos antes de actualizar.", vbExclamation, "ATENCIÓN"
            Exit Sub
        Else
            RsAux.Close
        End If
            
        On Error GoTo ErrBT
        cBase.BeginTrans    'COMIENZO LA TRANSACCION--------------------------
        On Error GoTo ErrET
        
        'Borro los precios A Regir para el articulo
        Cons = "Delete HistoriaPrecio " _
                & " Where HPrArticulo = " & Val(tArticulo.Tag) & " And HPrVigencia > GetDate()"
        cBase.Execute Cons
        
        cBase.CommitTrans    'FINALIZO LA TRANSACCION--------------------------
                
        LimpioCampos
        
        CargoDatosArticuloVigente
        If vsPrecio.Rows > 1 Then
            Botones True, True, False, False, False, Toolbar1, Me
        Else
            Botones True, False, False, False, False, Toolbar1, Me
        End If
        Screen.MousePointer = 0
    End If
    Exit Sub
    
ErrBT:
    clsGeneral.OcurrioError "No se pudo iniciar la transacción.", Err.Description
    Screen.MousePointer = vbDefault
    Exit Sub
ErrET:
    Resume ErrRoll
ErrRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "No se pudo eliminar la ficha de precios para el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionRefrescar()

    Screen.MousePointer = 11
    LimpioCampos
    
    If opVigencia(0).Value Then
        CargoDatosArticuloARegir
    Else
        CargoDatosArticuloVigente
    End If
    
    Screen.MousePointer = 0
    
End Sub

Sub AccionCancelar()

    Screen.MousePointer = 11
    Status.Panels(1).Text = ""
    vsPrecio.Editable = False
    If vsPrecio.Rows > 1 Then
        vsPrecio.Cell(flexcpBackColor, 1, 0, vsPrecio.Rows - 1, vsPrecio.Cols - 1) = vbWindowBackground
    End If
    
    If sModificar Then
    
        CargoDatosArticuloVigente
        Botones True, True, True, False, False, Toolbar1, Me
        
    Else
        
        CargoDatosArticuloARegir
        'Veo Si tiene precios vigentes para habilitar boton
        Cons = "Select * from PrecioVigente Where PViArticulo = " & Val(tArticulo.Tag)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            sVigente = True
        Else
            sVigente = False
        End If
        RsAux.Close
        
        Botones True, sVigente, True, False, False, Toolbar1, Me
        
        If vsPrecio.Rows = 1 Then CargoDatosArticuloVigente
        
    End If
    
    LimpioCampos
    DeshabilitoIngreso
    sNuevo = False
    sModificar = False
    
    Screen.MousePointer = 0
    
End Sub


Private Sub opVigencia_Click(Index As Integer)
    If Val(tArticulo.Tag) = 0 Then Exit Sub
    Select Case Index
        Case 0: CargoDatosArticuloVigente
        Case 1: CargoDatosArticuloARegir
    End Select
End Sub

Private Sub tArticulo_Change()
    If Val(tArticulo.Tag) > 0 Then
        tArticulo.Tag = ""
        vsPrecio.Rows = 1
        tPlan.Text = ""
        LimpioCampos
        Botones False, False, False, False, False, Toolbar1, Me
        opVigencia(0).Enabled = False
        opVigencia(1).Enabled = False
    End If
End Sub

Private Sub tArticulo_GotFocus()
On Error Resume Next
    With tArticulo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tArticulo_KeyPress(KeyAscii As Integer)
On Error GoTo errKP
Dim objLista As New clsListadeAyuda
    If KeyAscii = vbKeyReturn Then
        If Val(tArticulo.Tag) > 0 Then
            vsPrecio.SetFocus
        Else
            If IsNumeric(tArticulo.Text) Then
                CargoArticuloPorID Val(tArticulo.Text)
            Else
                Screen.MousePointer = 11
                Cons = "Select ArtCodigo as 'Código', ArtNombre as 'Nombre' From Articulo Where ArtNombre Like '" & clsGeneral.Replace(tArticulo.Text, " ", "%") & "%'"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If RsAux.EOF Then
                    RsAux.Close
                    Screen.MousePointer = 0
                    MsgBox "No existe un artículo con ese nombre.", vbInformation, "ATENCIÓN"
                Else
                    RsAux.MoveNext
                    If RsAux.EOF Then
                        RsAux.MoveFirst
                        tArticulo.Tag = RsAux(0)
                        RsAux.Close
                    Else
                        RsAux.Close
                        If objLista.ActivarAyuda(cBase, Cons, 5000, 0, "Lista de Artículos") > 0 Then
                            tArticulo.Tag = objLista.RetornoDatoSeleccionado(0)
                        End If
                    End If
                    Screen.MousePointer = 0
                    If Val(tArticulo.Tag) > 0 Then CargoArticuloPorID Val(tArticulo.Tag)
                End If
            End If
        End If
    End If
    Set objLista = Nothing
    Exit Sub
errKP:
    clsGeneral.OcurrioError "Ocurrió un error inesperado.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        Case "refrescar": AccionRefrescar
        Case "salir": Unload Me
    End Select

End Sub

Private Sub CargoDatosArticuloVigente()
Dim lAux As Long
    On Error GoTo errCargar
    
    opVigencia(0).Value = True
    vsPrecio.Rows = 1
    tPlan.Text = "": tPlan.Tag = "0"
    
    'CARGO DATOS DE TABLA Precios Vigentes-------------------
    Cons = "Select * " _
            & " From PrecioVigente, TipoCuota, Moneda, TipoPlan" _
            & " Where PViArticulo = " & tArticulo.Tag _
            & " And PViTipocuota = TCuCodigo" _
            & " And PViMoneda = MonCodigo And PViPlan = PlaCodigo" _
            & " Order by TCuOrden"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        'El plan debe ser de un ctdo y de pamonedapesos.
        'Y a su vez el precio este habilitado.
        If RsAux!TCuCodigo = paTipoCuotaContado And RsAux!MonCodigo = paMonedaFacturacion And RsAux!PViHabilitado Then
            
            tPlan.Text = Trim(RsAux!PlaNombre): tPlan.Tag = RsAux!PViPlan
            
        End If
        
        With vsPrecio
            .AddItem Trim(RsAux!TCuAbreviacion)
            lAux = RsAux!PViTipoCuota: .Cell(flexcpData, .Rows - 1, 0) = lAux
            lAux = RsAux!PViMoneda: .Cell(flexcpData, .Rows - 1, 1) = lAux
            lAux = RsAux!TCuCantidad: .Cell(flexcpData, .Rows - 1, 2) = lAux
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!MonSigno)
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!PViPrecio, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!PViPrecio / RsAux!TCuCantidad, "#,##0.00")
            .Cell(flexcpForeColor, .Rows - 1, 3) = ColorAjuste(.Cell(flexcpValue, .Rows - 1, 3))
            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!PViVigencia, "dd/mm/yy hh:mm")
            If RsAux!PViHabilitado Then
                .Cell(flexcpChecked, .Rows - 1, 5) = flexChecked
            Else
                .Cell(flexcpChecked, .Rows - 1, 5) = flexUnchecked
            End If
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al intentar cargar los datos del artículo.", Err.Description
    Screen.MousePointer = vbDefault
End Sub

Private Sub CargoDatosArticuloARegir()
On Error GoTo errCargar
Dim lAux As Long

    opVigencia(1).Value = True
    vsPrecio.Rows = 1
    'CARGO DATOS DE TABLA Precios Vigentes-------------------
    Cons = "Select PrecioARegir.*, TCuAbreviacion, TCuCantidad, MonSigno, PlaNombre, MonCodigo, TCuCodigo " _
            & " From PrecioARegir, TipoCuota, Moneda, TipoPlan" _
            & " Where PReArticulo = " & tArticulo.Tag _
            & " And PReTipocuota = TCuCodigo" _
            & " And PReMoneda = MonCodigo And PRePlan = PlaCodigo " _
            & " Order by TCuOrden"
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    tPlan.Text = "": tPlan.Tag = "0"
    Do While Not RsAux.EOF
        If RsAux!TCuCodigo = paTipoCuotaContado And RsAux!MonCodigo = paMonedaFacturacion Then
            tPlan.Text = Trim(RsAux!PlaNombre): tPlan.Tag = RsAux!PRePlan
        End If
        With vsPrecio
            .AddItem Trim(RsAux!TCuAbreviacion)
            lAux = RsAux!PReTipoCuota: .Cell(flexcpData, .Rows - 1, 0) = lAux
            lAux = RsAux!PReMoneda: .Cell(flexcpData, .Rows - 1, 1) = lAux
            lAux = RsAux!TCuCantidad: .Cell(flexcpData, .Rows - 1, 2) = lAux
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!MonSigno)
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!PRePrecio, "#,##0")
            .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!PRePrecio / RsAux!TCuCantidad, "#,##0")
            .Cell(flexcpForeColor, .Rows - 1, 3) = ColorAjuste(.Cell(flexcpValue, .Rows - 1, 3))
            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!PReVigencia, "dd/mm/yy hh:mm")
            If RsAux!PReHabilitado Then
                .Cell(flexcpText, .Rows - 1, 5) = flexChecked
            Else
                .Cell(flexcpText, .Rows - 1, 5) = flexUnchecked
            End If
            
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al intentar cargar los datos del artículo.", Err.Description
    Screen.MousePointer = vbDefault
End Sub

Function ValidoCampos()

    ValidoCampos = False
    
    If vsPrecio.Rows = 1 Then
        MsgBox "No se han ingresado datos.", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    
    If tPlan.Tag = "0" Then
        MsgBox "Se debe seleccionar el plan para la financiación del artículo.", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    
    ValidoCampos = True
    
End Function

Private Sub CargoCamposBDNuevo()
Dim lCont As Long
Dim rsOld As rdoResultset

    'Borro los Precios A Regir ------------------------------------------------------------
    Cons = "Delete HistoriaPrecio " _
           & " Where HPrArticulo = " & Val(tArticulo.Tag) & " And HPrVigencia > GetDate()"
    cBase.Execute Cons
    '------------------------------------------------------------------------------------------
    
    Cons = "Select * from HistoriaPrecio Where HPrArticulo = " & Val(tArticulo.Tag) _
        & " And HPrVigencia > GetDate()"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    With vsPrecio
    
        For lCont = 1 To .Rows - 1
        
            If vsPrecio.Cell(flexcpBackColor, lCont, 0) <> cteColorPV Then
                
                RsAux.AddNew
                RsAux!HPrArticulo = Val(tArticulo.Tag)
                RsAux!HPrTipoCuota = Val(.Cell(flexcpData, lCont, 0))
                RsAux!HPrMoneda = Val(.Cell(flexcpData, lCont, 1))
                RsAux!HPrVigencia = Format(.Cell(flexcpText, lCont, 4), "mm/dd/yyyy hh:nn:00")
                
                RsAux!HPrPlan = tPlan.Tag
                RsAux!HPrPrecio = CCur(.Cell(flexcpText, lCont, 2))
                        
                If .Cell(flexcpChecked, lCont, 5) = flexChecked Then
                    RsAux!HPrHabilitado = True
                Else
                    RsAux!HPrHabilitado = False
                End If
                RsAux.Update
                
            Else
                
                'Tengo que editar la ficha del precio vigente y deshabilitarlo.
                Cons = "Select * from HistoriaPrecio" _
                    & " Where HPrArticulo = " & Val(tArticulo.Tag) _
                    & " And HPrTipoCuota = " & Val(.Cell(flexcpData, lCont, 0)) _
                    & " And HPrMoneda = " & Val(.Cell(flexcpData, lCont, 1)) _
                    & " And HPrVigencia = '" & Format(.Cell(flexcpText, lCont, 4), "mm/dd/yyyy hh:nn:00") & "'" _
    
                Set rsOld = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not rsOld.EOF Then
                    rsOld.Edit
                
                    If .Cell(flexcpChecked, lCont, 5) = flexChecked Then
                        rsOld!HPrHabilitado = True
                    Else
                        rsOld!HPrHabilitado = False
                    End If
                    rsOld!HPrPrecio = CCur(.Cell(flexcpText, lCont, 2))
                    rsOld.Update
                End If
                rsOld.Close

            End If
        Next
    End With
    RsAux.Close

End Sub

Private Sub CargoCamposBDModificar()
Dim lCont As Long

    With vsPrecio
        For lCont = 1 To .Rows - 1
            
            Cons = "Select * from HistoriaPrecio" _
                    & " Where HPrArticulo = " & Val(tArticulo.Tag) _
                    & " And HPrTipoCuota = " & Val(.Cell(flexcpData, lCont, 0)) _
                    & " And HPrMoneda = " & Val(.Cell(flexcpData, lCont, 1)) _
                    & " And HPrVigencia = '" & Format(.Cell(flexcpText, lCont, 4), "mm/dd/yyyy hh:nn:00") & "'" _
    
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            RsAux.Edit
            
            If .Cell(flexcpChecked, lCont, 5) = flexChecked Then
                RsAux!HPrHabilitado = True
            Else
                RsAux!HPrHabilitado = False
            End If
            RsAux!HPrPrecio = CCur(.Cell(flexcpText, lCont, 2))
            RsAux.Update
            RsAux.Close
        Next
    End With
    
End Sub

Private Sub DeshabilitoIngreso()

    vsPrecio.Editable = False
    If vsPrecio.Rows > 1 Then vsPrecio.Cell(flexcpBackColor, 1, 0, vsPrecio.Rows - 1, vsPrecio.Cols - 1) = vbWindowBackground
    
    tArticulo.Enabled = True: tArticulo.BackColor = vbWindowBackground
    tPlan.BackColor = Inactivo
    cMoneda.BackColor = Inactivo
    cCuota.BackColor = Inactivo
    tValor.BackColor = Inactivo
    tVigencia.BackColor = Inactivo
       
    cCuota.Enabled = False
    tPlan.Enabled = False
    cMoneda.Enabled = False
    tValor.Enabled = False
    tVigencia.Enabled = False
    cHabilitado.Enabled = False
    
    If Val(tArticulo.Tag) > 0 Then
        opVigencia(0).Enabled = True
        opVigencia(1).Enabled = True
    Else
        opVigencia(0).Enabled = False
        opVigencia(1).Enabled = False
    End If
    
End Sub

Sub HabilitoIngreso()
    
    tArticulo.Enabled = False: tArticulo.BackColor = vbButtonFace
    If sNuevo Then
        tPlan.BackColor = Obligatorio
        cMoneda.BackColor = Obligatorio
        cCuota.BackColor = Obligatorio
        tValor.BackColor = Obligatorio
        tVigencia.BackColor = Obligatorio
                
        cMoneda.Enabled = True
        tPlan.Enabled = True
        cCuota.Enabled = True
        tValor.Enabled = True
        tVigencia.Enabled = True
        cHabilitado.Enabled = True
    End If
    Me.Refresh
    
End Sub

Private Sub tPlan_Change()
    tPlan.Tag = "0"
End Sub

Private Sub tPlan_GotFocus()
    With tPlan
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tPlan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(tPlan.Text) <> "" Then
            tPlan.Tag = "0"
            Screen.MousePointer = 11
            Cons = "Select Count(*) From TipoPlan Where PlaNombre Like '" & tPlan.Text & "%'"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
            If RsAux(0) = 0 Then
                RsAux.Close
                MsgBox "No existe un plan con ese descripción.", vbExclamation, "ATENCIÓN"
                Screen.MousePointer = 0
                Exit Sub
            ElseIf RsAux(0) = 1 Then
                Cons = "Select * From TipoPlan Where PlaNombre Like '" & tPlan.Text & "%'"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
                tPlan.Text = Trim(RsAux!PlaNombre)
                tPlan.Tag = RsAux!PlaCodigo
                RsAux.Close
                Foco tVigencia
            Else
                RsAux.Close
                Cons = "Select PlaCodigo, Nombre = PlaNombre From TipoPlan Where PlaNombre Like '" & tPlan.Text & "%'"
                Dim objAyuda As New clsListadeAyuda
                If objAyuda.ActivarAyuda(cBase, Cons, 4000, 1, "Ayuda de Planes") > 0 Then
                    Cons = "Select * From TipoPlan Where PlaCodigo = " & objAyuda.RetornoDatoSeleccionado(0)
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
                    tPlan.Text = Trim(RsAux!PlaNombre)
                    tPlan.Tag = RsAux!PlaCodigo
                    RsAux.Close
                    Foco tVigencia
                End If
                Set objAyuda = Nothing
            End If
        End If
        Screen.MousePointer = 0
    End If
End Sub

Private Sub tValor_GotFocus()

    tValor.SelStart = 0
    tValor.SelLength = Len(tValor.Text)
    
    Status.Panels(1).Text = "Ingrese el valor de la cuota en la moneda seleccionada."

End Sub


Private Sub tValor_LostFocus()

    tValor.Text = Format(tValor.Text, "#,##0.00")
    
End Sub

Private Sub tValor_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If cMoneda.ListIndex = -1 And sModificar Then
            MsgBox "Debe seleccionar una cuota de la lista y presionar ENTER para cargar la cuota a modificar.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
        If Trim(tValor.Text) <> "" Then cHabilitado.Value = 1
        cHabilitado.SetFocus
    End If

End Sub

Private Sub tVigencia_GotFocus()
    
    If Trim(tVigencia.Text) = "" Then
        If Weekday(Date + 1) = vbSunday Then
            tVigencia.Text = Format(Date + 2, "d/m/yyyy hh:mm")
        Else
            tVigencia.Text = Format(Date + 1, "d/m/yyyy hh:mm")
        End If
    End If
    tVigencia.SelStart = 0
    tVigencia.SelLength = Len(tVigencia.Text)
    Status.Panels(1).Text = "Ingrese la fecha de vigencia de los precios."
    
End Sub

Private Sub tVigencia_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        
        If UCase(Trim(tVigencia.Text)) = "A" Then
            tVigencia.Text = Format(DateAdd("n", 2, Now), "d/m/yyyy hh:nn")
        End If
        If Val(tPlan.Tag) <> 0 And IsDate(tVigencia.Text) And cMoneda.ListIndex = -1 Then
            BuscoCodigoEnCombo cMoneda, paMonedaFacturacion
            BuscoCodigoEnCombo cCuota, paTipoCuotaContado
        End If
        cMoneda.SetFocus
    End If
    
End Sub

Private Sub tVigencia_LostFocus()

    tVigencia.Text = Format(tVigencia.Text, "d/m/yyyy hh:mm")
    
End Sub

Private Sub CargoArticuloPorID(Optional lArtCodigo As Long = 0)
On Error GoTo errCA
    
    Screen.MousePointer = 11
    
    vsPrecio.Rows = 1
    opVigencia(0).Enabled = False: opVigencia(1).Enabled = False
    If lArtCodigo > 0 Then
        Cons = "Select * from Articulo where ArtCodigo = " & lArtCodigo
    Else
        Cons = "Select * FROM Articulo Where ArtID = " & Val(tArticulo.Tag)
    End If
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        tArticulo.Text = Trim(RsAux!ArtNombre)
        tArticulo.Tag = RsAux!ArtID
        aFecha = RsAux!ArtModificado
        RsAux.Close
        
        opVigencia(0).Enabled = True
        opVigencia(1).Enabled = True
        
        CargoDatosArticuloVigente
        If vsPrecio.Rows > 1 Then
            Botones True, True, True, False, False, Toolbar1, Me
        Else
            CargoDatosArticuloARegir
            If vsPrecio.Rows > 1 Then
                Botones True, False, True, False, False, Toolbar1, Me
            Else
                Botones True, False, False, False, False, Toolbar1, Me
            End If
        End If
    Else
        RsAux.Close
        Botones False, False, False, False, False, Toolbar1, Me
        LimpioCampos
    End If
    Screen.MousePointer = 0
    Exit Sub
errCA:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los artículos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub vsPrecio_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 3 Then
        vsPrecio.Cell(flexcpForeColor, Row, Col) = ColorAjuste(Val(vsPrecio.Cell(flexcpText, Row, Col)))
        vsPrecio.Cell(flexcpText, Row, 2) = Format(CCur(vsPrecio.Cell(flexcpText, Row, 3)) * CCur(vsPrecio.Cell(flexcpData, Row, 2)), "#,##0.00")
    End If
End Sub

Private Sub vsPrecio_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If Col <> 3 And Col <> 5 Then Cancel = True
    
End Sub

Private Sub vsPrecio_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    If Not sNuevo And Not sModificar Then Exit Sub
    
    'Elimino el registro---------------------------------------------------------
    If KeyCode = vbKeyDelete And vsPrecio.Rows > 1 And sNuevo Then
        If vsPrecio.Cell(flexcpBackColor, vsPrecio.Row) = cteColorPV Then
            If MsgBox("Al eliminar esta fila el precio quedará como vigente." & vbCrLf & _
                "¿Está seguro de que la cuota seleccionada quedará vigente?", vbQuestion + vbYesNo, "ATENCIÓN") <> vbYes Then Exit Sub
        
        Else
            Cons = "Select * From PrecioVigente, TipoCuota Where PViArticulo = " & tArticulo.Tag _
                    & " And PViHabilitado = 1 and PViMoneda = " & vsPrecio.Cell(flexcpData, vsPrecio.Row, 1) _
                    & " And PViTipoCuota = " & vsPrecio.Cell(flexcpData, vsPrecio.Row, 0) _
                    & " And PViTipoCuota = TCuCodigo"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
            If Not RsAux.EOF Then
                
                MsgBox "Se cargará el precio vigente para el tipo de cuota y la moneda.", vbInformation, "ATENCIÓN"
                
                With vsPrecio
                    .Cell(flexcpText, .Row, 2) = Format(RsAux!PViPrecio, "#,##0.00")
                    .Cell(flexcpText, .Row, 3) = Format(RsAux!PViPrecio / RsAux!TCuCantidad, "#,##0.00")
                    .Cell(flexcpForeColor, .Row, 3) = ColorAjuste(.Cell(flexcpValue, .Row, 3))
                    .Cell(flexcpText, .Row, 4) = Format(RsAux!PViVigencia, "dd/mm/yy hh:mm")
                
                    'x defecto lo presento deshabilitado.
                    .Cell(flexcpChecked, .Row, 5) = flexUnchecked
                    .Cell(flexcpBackColor, .Row, 0, , .Cols - 1) = cteColorPV
                End With
                RsAux.Close
                Exit Sub
            End If
            RsAux.Close
        End If
        vsPrecio.RemoveItem vsPrecio.Row
        Exit Sub
    End If

    'Edito el registro para modificar------------------------------------------
'    If KeyCode = vbKeyReturn And vsPrecio.Row > 0 Then
 '       Exit Sub
        
        'LO INHABILITE 24-6-2002 ya que dejo editar la grilla.
        
'        BuscoCodigoEnCombo cCuota, vsPrecio.Cell(flexcpData, vsPrecio.Row, 0)
 '       BuscoCodigoEnCombo cMoneda, vsPrecio.Cell(flexcpData, vsPrecio.Row, 1)
  '      tValor.Text = vsPrecio.Cell(flexcpText, vsPrecio.Row, 3)
        
   '     tVigencia.Text = vsPrecio.Cell(flexcpText, vsPrecio.Row, 4)
    '    If vsPrecio.Cell(flexcpChecked, vsPrecio.Row, 5) = flexChecked Then
     '       cHabilitado.Value = 1
      '  Else
       '     cHabilitado.Value = 0
        'End If
        'If sNuevo Then
         '   vsPrecio.RemoveItem vsPrecio.Row
          '  tValor.Enabled = True: tValor.BackColor = Obligatorio
           ' tValor.SetFocus
        'Else
         '   tValor.BackColor = Obligatorio
          '  tValor.Enabled = True
           ' cHabilitado.Enabled = True
            'tValor.SetFocus
        'End If
    'End If
    '-----------------------------------------------------------------------------
End Sub

Private Sub vsPrecio_RowColChange()
    
    If Not sNuevo And Not sModificar Then
        Status.Panels(1).Text = ""
    Else
        If sNuevo Then
            Status.Panels(1).Text = "Para eliminar presione [Supr]"
        Else
            Status.Panels(1).Text = "Se podrá editar el precio de la cuota o deshabilitar una cuota."
        End If
    End If
    
End Sub

Public Function ColorAjuste(ByVal Importe As Currency) As Long
On Error Resume Next
    ColorAjuste = vsPrecio.ForeColor
    If Importe = 0 Then ColorAjuste = vbRed: Exit Function
    If (Right(Importe, 2) / Importe) < 0.008 Then ColorAjuste = vbRed
    
End Function

Private Sub vsPrecio_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If sModificar Then
        If Col = 3 Then
            If Not IsNumeric(vsPrecio.EditText) Then
                Cancel = True
                MsgBox "El precio de la cuota debe ser un valor numérico.", vbExclamation, "ATENCIÓN"
            Else
                vsPrecio.EditText = Format(vsPrecio.EditText, "#,##0.00")
            End If
        End If
    End If
    
End Sub

Private Sub ControlPrecioNuevoContraVigentes()
Dim iCant As Integer
Dim lAux As Long

    'Levanto los precios vigentes.
    'Si tengo un tipo de cuota que no este ingresado con el plan
    'nuevo ---> lo presento pero en color rojo.
    
    If vsPrecio.Rows = 1 Then Exit Sub
        
    Cons = "Select * " _
            & " From PrecioVigente, TipoCuota, Moneda, TipoPlan" _
            & " Where PViArticulo = " & tArticulo.Tag _
            & " And PViHabilitado = 1 And PViTipocuota = TCuCodigo" _
            & " And PViMoneda = MonCodigo And PViPlan = PlaCodigo" _
            & " Order by TCuOrden"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        
        'Para este precio recorro la lista para ver si lo encuentro.
        
        iCant = 0
        For iCant = 1 To vsPrecio.Rows - 1
            If CCur(vsPrecio.Cell(flexcpData, iCant, 0)) = RsAux!PViTipoCuota And _
                CCur(vsPrecio.Cell(flexcpData, iCant, 1)) = RsAux!PViMoneda Then
                    iCant = -1
                    Exit For
            End If
        Next
        
        If iCant <> -1 Then
            With vsPrecio
                .AddItem Trim(RsAux!TCuAbreviacion)
                lAux = RsAux!PViTipoCuota: .Cell(flexcpData, .Rows - 1, 0) = lAux
                lAux = RsAux!PViMoneda: .Cell(flexcpData, .Rows - 1, 1) = lAux
                lAux = RsAux!TCuCantidad: .Cell(flexcpData, .Rows - 1, 2) = lAux
                .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!MonSigno)
                .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!PViPrecio, "#,##0.00")
                .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!PViPrecio / RsAux!TCuCantidad, "#,##0.00")
                .Cell(flexcpForeColor, .Rows - 1, 3) = ColorAjuste(.Cell(flexcpValue, .Rows - 1, 3))
                .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!PViVigencia, "dd/mm/yy hh:mm")
                
                'x defecto lo presento deshabilitado.
                .Cell(flexcpChecked, .Rows - 1, 5) = flexUnchecked
                
                .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = cteColorPV
            End With
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close

End Sub
