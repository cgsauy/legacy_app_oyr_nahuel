VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACOMBO.OCX"
Begin VB.Form MaPrecioArticulo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Precios de Artículos"
   ClientHeight    =   4455
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
   Icon            =   "MaPrecioArticulo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6990
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   11
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   500
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "refrescar"
            Object.ToolTipText     =   "Refrescar"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   3550
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin AACombo99.AACombo cMoneda 
      Height          =   315
      Left            =   720
      TabIndex        =   5
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
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox tValor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   8
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox tVigencia 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   4320
      MaxLength       =   16
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.CheckBox cHabilitado 
      Alignment       =   1  'Right Justify
      Caption         =   "&En Uso:"
      Height          =   255
      Left            =   6000
      TabIndex        =   9
      Top             =   1320
      Width           =   850
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   4200
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   8625
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   3625
            MinWidth        =   2
            Text            =   "Ctr+H - Historia de Precios "
            TextSave        =   "Ctr+H - Historia de Precios "
            Key             =   ""
            Object.Tag             =   ""
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
   Begin ComctlLib.ListView lPrecio 
      Height          =   2415
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Tipo de Cuota"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Moneda"
         Object.Width           =   441
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Precio"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Valor Cuota"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Vigencia al"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "En Uso"
         Object.Width           =   459
      EndProperty
   End
   Begin AACombo99.AACombo cCuota 
      Height          =   315
      Left            =   1560
      TabIndex        =   6
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Lista"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lTipo 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "Vigentes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   16
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "&Cuotas:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "V&alor Cuota:"
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "&Plan:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "&Vigencia:"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lArticulo 
      BackColor       =   &H00C00000&
      Caption         =   "Artículo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   15
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   " Artículo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   855
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7080
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaPrecioArticulo.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaPrecioArticulo.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaPrecioArticulo.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaPrecioArticulo.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaPrecioArticulo.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaPrecioArticulo.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaPrecioArticulo.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaPrecioArticulo.frx":0DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaPrecioArticulo.frx":10E2
            Key             =   ""
         EndProperty
      EndProperty
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
Attribute VB_Name = "MaPrecioArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Cambios
'25-10-2001
'    En distribuir cuotas cambiamos consulta para que no cargue los coeficientes = 1 y las Tipo cuotas que no estén habilitadas.

Private sNuevo As Boolean, sModificar As Boolean, sAyuda As Boolean
Private sVigente As Boolean             'Indica si hay precios vigentes ingresados (para Hab/Des botones)

Private RsArticulo As rdoResultset
Private iSeleccionado As Long

Private aFecha As Date        'Para controlar modificaciones multiusuario

Public Property Get pSeleccionado() As Long
    pSeleccionado = iSeleccionado
End Property
Public Property Let pSeleccionado(Codigo As Long)
    iSeleccionado = Codigo
End Property

Private Sub DistribuirCuotas(Plan As Integer, Moneda As Integer, Contado As Currency)

Dim sCargueElContado As Boolean     'Para saber si el contado fue ingresado (por si no pertenece al plan)

    sCargueElContado = False
    'Saco los datos de los coeficientes con = moneda a la del precio y sin entrega (TCuVencimientoE = NULL)
    Cons = "Select * from Coeficiente, TipoCuota" _
        & " Where CoePlan =  " & Plan _
        & " And CoeTipoCuota = TCuCodigo" _
        & " And CoeMoneda = " & Moneda _
        & " And TCuVencimientoE Is NULL And TCuDeshabilitado Is Null " _
        & " And CoeCoeficiente <> 1 Order by TCuOrden"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
    
        If RsAux!TCuCodigo = paTipoCuotaContado Then sCargueElContado = True
        
        'Verifico si hay ingresado un precio para la misma moneda---
        For Each itmx In lPrecio.ListItems
            If Mid(itmx.Key, 1, InStr(itmx.Key, "V") - 1) = RsAux!CoeTipoCuota & "M" & Moneda Then
                lPrecio.ListItems.Remove itmx.Index
                Exit For
            End If
        Next
        '----------------------------------------------------------------------
        
        'TipoCuota "M" Moneda "V" Venvimiento
        Set itmx = lPrecio.ListItems.Add(, RsAux!TCuCodigo & "M" & RsAux!CoeMoneda & "V" & Format(tVigencia.Text, "dd/mm/yyyy hh:mm"), Trim(RsAux!TCuAbreviacion))
        itmx.SubItems(1) = cMoneda.Text
        'Valor de la Cuota
        itmx.SubItems(3) = Format((Contado * RsAux!CoeCoeficiente) / RsAux!TCuCantidad, "#,##0") & ".00"
        'Precio Total
        itmx.SubItems(2) = Format(CCur(itmx.SubItems(3)) * RsAux!TCuCantidad, "#,##0") & ".00"
        
        itmx.SubItems(4) = Format(tVigencia.Text, "Ddd d/mm/yy hh:mm")
        If cHabilitado.Value = 1 Then
            itmx.SubItems(5) = "Si"
        Else
            itmx.SubItems(5) = "No"
        End If
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If Not sCargueElContado Then
        'Es un Contado que no pertenece al plan (la excepcion a la regla)
        AgregoCuota cCuota.ItemData(cCuota.ListIndex), Moneda
    End If
    
End Sub

Private Sub AgregoCuota(Cuota As Integer, Moneda As Integer)

Dim aCantidad As Integer

    Cons = "Select * From TipoCuota Where TCuCodigo = " & Cuota
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    aCantidad = RsAux!TCuCantidad
    RsAux.Close
    
    'TipoCuota "M" Moneda "V" Venvimiento
    Set itmx = lPrecio.ListItems.Add(, Cuota & "M" & Moneda & "V" & Format(tVigencia.Text, "dd/mm/yyyy hh:mm"), Trim(cCuota.Text))
    itmx.SubItems(1) = cMoneda.Text
    'Valor de la Cuota
    itmx.SubItems(3) = Format(tValor.Text, "#,##0.00")
    'Precio Total
    itmx.SubItems(2) = Format(CCur(itmx.SubItems(3)) * aCantidad, "#,##0.00")
    
    itmx.SubItems(4) = Format(tVigencia.Text, "Ddd d/mm/yy hh:mm")
    If cHabilitado.Value = 1 Then
        itmx.SubItems(5) = "Si"
    Else
        itmx.SubItems(5) = "No"
    End If
    
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
                        
            'Verifico si hay ingresado un precio para la misma moneda---
            For Each itmx In lPrecio.ListItems
                If Mid(itmx.Key, 1, InStr(itmx.Key, "V") - 1) = cCuota.ItemData(cCuota.ListIndex) & "M" & cMoneda.ItemData(cMoneda.ListIndex) Then
                    MsgBox "Existe un precio para la moneda y tipo de cuota seleccionados, para modificarlo edítelo.", vbExclamation, "ATENCIÓN"
                    Exit Sub
                End If
            Next
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
            
            'Valor de la Cuota
            lPrecio.SelectedItem.SubItems(3) = Format(CCur(tValor.Text), "#,##0") & ".00"
            'Precio Total
            lPrecio.SelectedItem.SubItems(2) = Format(CCur(lPrecio.SelectedItem.SubItems(3)) * lPrecio.SelectedItem.Tag, "#,##0") & ".00"
            
            If cHabilitado.Value = 0 Then
                lPrecio.SelectedItem.SubItems(5) = "No"
            Else
                lPrecio.SelectedItem.SubItems(5) = "Si"
            End If
            LimpioCampos
            lPrecio.SetFocus
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
    If Not sAyuda Then
        RsArticulo.Requery
    End If
    
    Screen.MousePointer = 0
    DoEvents

End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    
    ObtengoSeteoForm Me, 500, 500
    SetearLView lvValores.Grilla Or lvValores.FullRow, lPrecio
    
    sNuevo = False: sModificar = False: sAyuda = False
    
    
    DeshabilitoIngreso
    
    'Cargo las MONEDAS
    Cons = "Select MonCodigo, MonSigno From Moneda Order by MonSigno"
    CargoCombo Cons, cMoneda, ""
    
    'Cargo los Tipos de Cuotas
    Cons = "Select TCuCodigo, TCuAbreviacion from TipoCuota" _
            & " Where TCuVencimientoE = NULL"
    CargoCombo Cons, cCuota, ""

    Cons = "Select * FROM Articulo Where ArtID = " & iSeleccionado
    Set RsArticulo = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsArticulo.EOF Then
        CargoDatosArticuloVigente
        If lPrecio.ListItems.Count > 0 Then
            Botones True, True, True, False, False, Toolbar1, Me
        Else
            CargoDatosArticuloARegir
            If lPrecio.ListItems.Count > 0 Then
                Botones True, False, True, False, False, Toolbar1, Me
            Else
                Botones True, False, False, False, False, Toolbar1, Me
            End If
        End If
    Else
        MsgBox "El artículo seleccionado ha sido eliminado o no se puede acceder al registro.", vbExclamation, "ERROR"
        Botones False, False, False, False, False, Toolbar1, Me
    End If
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al ingresar al formulario."
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.Panels(1).Text = ""
    
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
    If Not gCerrarConexion Then
        RsArticulo.Close
        Forms(Forms.Count - 2).SetFocus
    Else
        RsArticulo.Close
        CierroConexion
        Set clsGeneral = Nothing
        Set miConexion = Nothing
        Forms(Forms.Count - 2).SetFocus
        End
    End If
    
End Sub


Private Sub Label17_Click()
    Foco tVigencia
End Sub

Private Sub Label25_Click()
    Foco tPlan
End Sub

Private Sub Label26_Click()
    Foco tValor
End Sub

Private Sub Label27_Click()
    Foco cMoneda
End Sub

Private Sub lPrecio_GotFocus()

    Status.Panels(1).Text = "Lista de precios ingresados"
    
End Sub

Private Sub lPrecio_KeyDown(KeyCode As Integer, Shift As Integer)

    If Not sNuevo And Not sModificar Then Exit Sub
    
    'Elimino el registro---------------------------------------------------------
    If KeyCode = vbKeyDelete And lPrecio.ListItems.Count > 0 And sNuevo Then
        lPrecio.ListItems.Remove (lPrecio.SelectedItem.Key)
    End If

    'Edito el registro para modificar------------------------------------------
    If KeyCode = vbKeyReturn And lPrecio.ListItems.Count > 0 Then
        BuscoCodigoEnCombo cCuota, Mid(lPrecio.SelectedItem.Key, 1, InStr(lPrecio.SelectedItem.Key, "M") - 1)
        BuscoCodigoEnCombo cMoneda, Mid(lPrecio.SelectedItem.Key, InStr(lPrecio.SelectedItem.Key, "M") + 1, InStr(lPrecio.SelectedItem.Key, "V") - InStr(lPrecio.SelectedItem.Key, "M") - 1)
        tValor.Text = lPrecio.SelectedItem.SubItems(3)
        tVigencia.Text = Format(Mid(lPrecio.SelectedItem.Key, InStr(lPrecio.SelectedItem.Key, "V") + 1, Len(lPrecio.SelectedItem.Key)), "d/m/yyyy hh:mm")
        If UCase(lPrecio.SelectedItem.SubItems(5)) = "SI" Then
            cHabilitado.Value = 1
        Else
            cHabilitado.Value = 0
        End If
        If sNuevo Then
            lPrecio.ListItems.Remove (lPrecio.SelectedItem.Key)
            tValor.SetFocus
        Else
            If tValor.Enabled Then tValor.SetFocus Else cHabilitado.SetFocus
        End If
    End If
    '-----------------------------------------------------------------------------
    
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
           & " And HPrArticulo = " & iSeleccionado & " And HPrPlan = PlaCodigo " _
           & "Order by HPrVigencia DESC"
           
    Dim LiAyuda As New clsListadeAyuda
    LiAyuda.ActivoListaAyudaSQL cBase, Cons
    Set LiAyuda = Nothing
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
    
    lTipo.Caption = "A Regir"
    Foco tPlan
    Screen.MousePointer = 0
    Exit Sub
ErrAN:
    clsGeneral.OcurrioError "Ocurrio un error inesperado.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub LimpioCampos()
    
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
    CargoDatosArticuloVigente
    Screen.MousePointer = 0
    
    lTipo.Caption = "Vigentes"

End Sub

Sub AccionGrabar()

Dim aCodigo As Long

    If Not ValidoCampos Then Exit Sub
        
    If MsgBox("Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then
        Screen.MousePointer = 11
        
        RsArticulo.Requery
        If aFecha <> RsArticulo!ArtModificado Then
            Screen.MousePointer = 0
            MsgBox "La ficha ha sido modificada por otro usuario. Verifique los datos antes de grabar.", vbExclamation, "ATENCIÓN"
            AccionCancelar
            Exit Sub
        End If
        
        If sNuevo Then      'NUEVO ARTICULO
            On Error GoTo ErrBT
            cBase.BeginTrans    'COMIENZO LA TRANSACCION--------------------------
            On Error GoTo ErrET
            
            RsArticulo.Requery
            
            'Actualizo la fecha de modificacion del artículo
            RsArticulo.Edit
            RsArticulo!ArtModificado = Format(Now, sqlFormatoFH)
            RsArticulo.Update
            
            'Cargo los datos en HistoriaPrecio
            CargoCamposBDNuevo
            
            cBase.CommitTrans   'FIN DE LA TRANSACCION-------------------------------
            sNuevo = False
            
            RsArticulo.Requery
            
            'Veo Si tiene precios vigentes para habilitar boton--------------
            Cons = "Select * from PrecioVigente Where PViArticulo = " & iSeleccionado
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                sVigente = True
            Else
                sVigente = False
            End If
            RsAux.Close '--------------------------------------------------------
        
            Botones True, sVigente, True, False, False, Toolbar1, Me
            
        Else                      'MODIFICACION DE LOS DATOS
            
            On Error GoTo ErrBT
            cBase.BeginTrans    'COMIENZO LA TRANSACCION--------------------------
            On Error GoTo ErrET
            
            aCodigo = RsArticulo!ArtID
            RsArticulo.Requery
            
            'Actualizo la fecha de modificacion del artículo
            RsArticulo.Edit
            RsArticulo!ArtModificado = Format(Now, sqlFormatoFH)
            RsArticulo.Update
            
            'Cargo los datos en HistoriaPrecio
            CargoCamposBDModificar
            
            sModificar = False
            cBase.CommitTrans   'FIN DE LA TRANSACCION-------------------------------
            RsArticulo.Requery
            Botones True, True, True, False, False, Toolbar1, Me
        End If
        
        DeshabilitoIngreso
        LimpioCampos
        Screen.MousePointer = 0
    End If
    Exit Sub
    
ErrBT:
    clsGeneral.OcurrioError "No se pudo iniciar la transacción.", Err.Description
    RsArticulo.Requery
    Screen.MousePointer = vbDefault
    Exit Sub
ErrET:
    Resume ErrRoll
ErrRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "No se pudo almacenar la información, reintente."
    RsArticulo.Requery
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Sub AccionEliminar()

    Screen.MousePointer = 11
    'Veo Si tiene precios A REGIR para eliminar
    Cons = "Select * from PrecioARegir Where PReArticulo = " & iSeleccionado
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        Screen.MousePointer = 0
        MsgBox "No hay precios A Regir para eliminar.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    RsAux.Close
    
    If MsgBox("Confrima eliminar la lista de precios a regir.", vbQuestion + vbYesNo, "ELIMINAR") = vbYes Then
        Screen.MousePointer = 11
        RsArticulo.Requery
        If aFecha <> RsArticulo!ArtModificado Then
            Screen.MousePointer = 0
            MsgBox "La ficha ha sido modificada por otro usuario. Verifique los datos antes de actualizar.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
            
        On Error GoTo ErrBT
        cBase.BeginTrans    'COMIENZO LA TRANSACCION--------------------------
        On Error GoTo ErrET
        
        RsArticulo.Requery
        
        'Borro los precios A Regir para el articulo
        Cons = "Delete HistoriaPrecio " _
                & " Where HPrArticulo = " & iSeleccionado _
                & " And HPrVigencia > GetDate()"
                
        cBase.Execute Cons
        
        cBase.CommitTrans    'FINALIZO LA TRANSACCION--------------------------
        
        RsArticulo.Requery
                
        LimpioCampos
        lPrecio.ListItems.Clear
        
        CargoDatosArticuloVigente
        If lPrecio.ListItems.Count > 0 Then
            Botones True, True, False, False, False, Toolbar1, Me
        Else
            Botones True, False, False, False, False, Toolbar1, Me
        End If
        Screen.MousePointer = 0
    End If
    Exit Sub
    
ErrBT:
    clsGeneral.OcurrioError "No se pudo iniciar la transacción."
    RsArticulo.Requery
    Screen.MousePointer = vbDefault
    Exit Sub
ErrET:
    Resume ErrRoll
ErrRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "No se pudo eliminar la ficha de precios para el artículo."
    RsArticulo.Requery
    Screen.MousePointer = 0
End Sub

Private Sub AccionRefrescar()

    Screen.MousePointer = 11
    LimpioCampos
    
    If Trim(lTipo.Caption) = "A Regir" Then
        CargoDatosArticuloARegir
    Else
        CargoDatosArticuloVigente
    End If
    
    Screen.MousePointer = 0
    
End Sub

Sub AccionCancelar()

    Screen.MousePointer = 11
     
    If sModificar Then
        CargoDatosArticuloVigente
        Botones True, True, True, False, False, Toolbar1, Me
        
    Else
        CargoDatosArticuloARegir
        'Veo Si tiene precios vigentes para habilitar boton
        Cons = "Select * from PrecioVigente Where PViArticulo = " & iSeleccionado
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            sVigente = True
        Else
            sVigente = False
        End If
        RsAux.Close
        
        Botones True, sVigente, True, False, False, Toolbar1, Me
        
        If lPrecio.ListItems.Count = 0 Then
            CargoDatosArticuloVigente
        End If
    End If
    
    LimpioCampos
    DeshabilitoIngreso
    sNuevo = False
    sModificar = False
    
    Screen.MousePointer = 0
    
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        
        Case "nuevo"
            AccionNuevo
        
        Case "modificar"
            AccionModificar
        
        Case "eliminar"
            AccionEliminar
        
        Case "grabar"
            AccionGrabar
        
        Case "cancelar"
            AccionCancelar
        
        Case "refrescar"
            AccionRefrescar
            
        Case "salir"
            Unload Me
            
    End Select

End Sub

Private Sub CargoDatosArticuloVigente()

    On Error GoTo errCargar
    lPrecio.ListItems.Clear
    tPlan.Text = "": tPlan.Tag = "0"
    
    aFecha = RsArticulo!ArtModificado
    
    'CARGO DATOS DE TABLA Articulo -------------------
    lArticulo = Trim(RsArticulo!ArtNombre)
    
    'CARGO DATOS DE TABLA Precios Vigentes-------------------
    Cons = "Select PrecioVigente.*, TCuAbreviacion, TCuCantidad, MonSigno, PlaNombre " _
            & " From PrecioVigente, TipoCuota, Moneda, TipoPlan" _
            & " Where PViArticulo = " & RsArticulo!ArtID _
            & " And PViTipocuota = TCuCodigo" _
            & " And PViMoneda = MonCodigo And PViPlan = PlaCodigo" _
            & " Order by TCuOrden"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        tPlan.Text = Trim(RsAux!PlaNombre): tPlan.Tag = RsAux!PViPlan
        Set itmx = lPrecio.ListItems.Add(, RsAux!PViTipoCuota & "M" & RsAux!PViMoneda & "V" & RsAux!PViVigencia, Trim(RsAux!TCuAbreviacion))
        itmx.Tag = RsAux!TCuCantidad
        itmx.SubItems(1) = Trim(RsAux!MonSigno)
        itmx.SubItems(2) = Format(RsAux!PViPrecio, "#,##0.00")
        itmx.SubItems(3) = Format(RsAux!PViPrecio / RsAux!TCuCantidad, "#,##0.00")
        itmx.SubItems(4) = Format(RsAux!PViVigencia, "Ddd d/mm/yy hh:mm")
        If RsAux!PViHabilitado Then
            itmx.SubItems(5) = "Si"
        Else
            itmx.SubItems(5) = "No"
        End If
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    If lPrecio.ListItems.Count > 0 Then lTipo.Caption = "Vigentes"
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al intentar cargar los datos del artículo.", Err.Description
    Screen.MousePointer = vbDefault
End Sub

Private Sub CargoDatosArticuloARegir()

    On Error GoTo errCargar
    lPrecio.ListItems.Clear
    
    aFecha = RsArticulo!ArtModificado
    
    'CARGO DATOS DE TABLA Articulo -------------------
    lArticulo = Trim(RsArticulo!ArtNombre)
    
    'CARGO DATOS DE TABLA Precios Vigentes-------------------
    Cons = "Select PrecioARegir.*, TCuAbreviacion, TCuCantidad, MonSigno, PlaNombre " _
            & " From PrecioARegir, TipoCuota, Moneda, TipoPlan" _
            & " Where PReArticulo = " & RsArticulo!ArtID _
            & " And PReTipocuota = TCuCodigo" _
            & " And PReMoneda = MonCodigo And PRePlan = PlaCodigo " _
            & " Order by TCuOrden"
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    tPlan.Text = "": tPlan.Tag = "0"
    Do While Not RsAux.EOF
        tPlan.Text = Trim(RsAux!PlaNombre): tPlan.Tag = RsAux!PRePlan
        Set itmx = lPrecio.ListItems.Add(, RsAux!PReTipoCuota & "M" & RsAux!PReMoneda & "V" & RsAux!PReVigencia, Trim(RsAux!TCuAbreviacion))
        itmx.SubItems(1) = Trim(RsAux!MonSigno)
        itmx.SubItems(2) = Format(RsAux!PRePrecio, "#,##0.00")
        itmx.SubItems(3) = Format(RsAux!PRePrecio / RsAux!TCuCantidad, "#,##0.00")
        itmx.SubItems(4) = Format(RsAux!PReVigencia, "Ddd d/mm/yy hh:mm")
        If RsAux!PReHabilitado Then
            itmx.SubItems(5) = "Si"
        Else
            itmx.SubItems(5) = "No"
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    If lPrecio.ListItems.Count > 0 Then lTipo.Caption = "A Regir"
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al intentar cargar los datos del artículo.", Err.Description
    Screen.MousePointer = vbDefault
End Sub

Function ValidoCampos()

    ValidoCampos = False
    
    If lPrecio.ListItems.Count = 0 Then
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

    'Borro los Precios A Regir ------------------------------------------------------------
    Cons = "Delete HistoriaPrecio " _
           & " Where HPrArticulo = " & iSeleccionado _
           & " And HPrVigencia > GetDate()"
    cBase.Execute Cons
    '------------------------------------------------------------------------------------------
    
    Cons = "Select * from HistoriaPrecio Where HPrArticulo = " & iSeleccionado
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    For Each itmx In lPrecio.ListItems
                
        RsAux.AddNew
        RsAux!HPrArticulo = iSeleccionado
        RsAux!HPrTipoCuota = Mid(itmx.Key, 1, InStr(itmx.Key, "M") - 1)
        RsAux!HPrMoneda = Mid(itmx.Key, InStr(itmx.Key, "M") + 1, InStr(itmx.Key, "V") - InStr(itmx.Key, "M") - 1)
        RsAux!HPrVigencia = Format(Mid(itmx.Key, InStr(itmx.Key, "V") + 1, Len(itmx.Key)), sqlFormatoFH)
        
        RsAux!HPrPlan = tPlan.Tag
        RsAux!HPrPrecio = CCur(itmx.SubItems(2))
                
        If UCase(itmx.SubItems(5)) = "SI" Then
            RsAux!HPrHabilitado = True
        Else
            RsAux!HPrHabilitado = False
        End If
        
        RsAux.Update
        
    Next
    RsAux.Close

End Sub

Private Sub CargoCamposBDModificar()

    For Each itmx In lPrecio.ListItems
        
        Cons = "Select * from HistoriaPrecio" _
                & " Where HPrArticulo = " & iSeleccionado _
                & " And HPrTipoCuota = " & Mid(itmx.Key, 1, InStr(itmx.Key, "M") - 1) _
                & " And HPrMoneda = " & Mid(itmx.Key, InStr(itmx.Key, "M") + 1, InStr(itmx.Key, "V") - InStr(itmx.Key, "M") - 1) _
                & " And HPrVigencia = '" & Format(Mid(itmx.Key, InStr(itmx.Key, "V") + 1, Len(itmx.Key)), sqlFormatoFH) & "'" _

        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.Edit
        
        If UCase(itmx.SubItems(5)) = "SI" Then
            RsAux!HPrHabilitado = True
        Else
            RsAux!HPrHabilitado = False
        End If
        RsAux!HPrPrecio = CCur(itmx.SubItems(2))
        
        RsAux.Update
        RsAux.Close
    Next
    
End Sub

Private Sub DeshabilitoIngreso()

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
    
End Sub

Sub HabilitoIngreso()
    
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
    Else
        tValor.BackColor = Obligatorio
        tValor.Enabled = True
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
        If WeekDay(Date + 1) = vbSunday Then
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
        If UCase(Trim(tVigencia.Text)) = "A" Then tVigencia.Text = Format(Now, "d/m/yyyy hh") & ":" & Minute(Now) + 2
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

