VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form DeEntrega 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Devolución de Entregas"
   ClientHeight    =   4875
   ClientLeft      =   2490
   ClientTop       =   2325
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DeEntrega.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6840
   Begin VB.TextBox tComentario 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1200
      MaxLength       =   100
      TabIndex        =   23
      Top             =   4500
      Width           =   5535
   End
   Begin VB.TextBox tUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5400
      MaxLength       =   2
      TabIndex        =   10
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton bCancelar 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   5880
      TabIndex        =   14
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton bDFactura 
      Caption         =   "&Detalle de Factura"
      Height          =   315
      Left            =   120
      TabIndex        =   22
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton bFinalizar 
      Caption         =   "Finali&zar"
      Height          =   315
      Left            =   4920
      TabIndex        =   13
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox tArticulo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3120
      MaxLength       =   30
      TabIndex        =   12
      Top             =   2280
      Width           =   1695
   End
   Begin VB.ComboBox cSucursal 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox tSerie 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox tNumero 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4440
      MaxLength       =   6
      TabIndex        =   8
      Top             =   960
      Width           =   855
   End
   Begin VB.ComboBox cTipo 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox tCBarra 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   320
      Width           =   2415
   End
   Begin ComctlLib.ListView lArticulo 
      Height          =   1575
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
         Text            =   "Código"
         Object.Width           =   1500
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Articulo"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "A Devolver"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Devueltos"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   "BarCode"
         Text            =   ""
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   "TipoArticulo"
         Text            =   ""
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Co&mentarios:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4545
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario"
      Height          =   255
      Left            =   5400
      TabIndex        =   9
      Top             =   720
      Width           =   615
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5940
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DeEntrega.frx":030A
            Key             =   "si"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DeEntrega.frx":0624
            Key             =   "no"
         EndProperty
      EndProperty
   End
   Begin VB.Label lAEntregar 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4200
      Width           =   6600
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "DATOS DEL CLIENTE..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lTitular 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "N/D"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   1920
      UseMnemonic     =   0   'False
      Width           =   5535
   End
   Begin VB.Label lCiRuc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "N/D"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1080
      TabIndex        =   17
      Top             =   1680
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "CI/RUC:"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "DE&VOLUCIONES:"
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   2325
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Sucursal"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "&Número"
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "&Tipo"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "&Factura o Remito"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      Height          =   840
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1395
      Width           =   6615
   End
End
Attribute VB_Name = "DeEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Rojo = &HC0C0FF

Dim gDocumento As Long
Dim gTipo As Integer
Dim gFechaDocumento As Date
Dim gFechaEmision As Date

Dim rsXX As rdoResultset
Dim Fletes As String

'------------------------------------------------------------------
Private Type typNroSerie
    Articulo As Long
    NroSerie As String
End Type

Dim arrNroSerie() As typNroSerie
'------------------------------------------------------------------

Private Sub bCancelar_Click()
    LimpioParaNuevaEntrega
End Sub

Private Sub bDFactura_Click()
    If gDocumento = 0 Then Exit Sub
    EjecutarApp pathApp & "\Detalle de factura", CStr(gDocumento)
End Sub

Private Sub bFinalizar_Click()
    AccionGrabar
End Sub

Private Sub cSucursal_Change()
    Selecciono cSucursal, cSucursal.Text, gTecla
End Sub

Private Sub cSucursal_GotFocus()
    cSucursal.SelStart = 0
    cSucursal.SelLength = Len(cSucursal.Text)
End Sub

Private Sub cSucursal_KeyDown(KeyCode As Integer, Shift As Integer)
    gTecla = KeyCode
    gIndice = cSucursal.ListIndex
End Sub

Private Sub cSucursal_KeyPress(KeyAscii As Integer)
    cSucursal.ListIndex = gIndice
    If KeyAscii = vbKeyReturn And cSucursal.ListIndex <> -1 Then Foco tSerie
End Sub

Private Sub cSucursal_KeyUp(KeyCode As Integer, Shift As Integer)
    ComboKeyUp cSucursal
End Sub

Private Sub cSucursal_LostFocus()
    gIndice = -1
    cSucursal.SelLength = 0
End Sub

Private Sub cTipo_Change()
    Selecciono cTipo, cTipo.Text, gTecla
End Sub

Private Sub cTipo_GotFocus()
    cTipo.SelStart = 0
    cTipo.SelLength = Len(cTipo.Text)
End Sub

Private Sub cTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    gTecla = KeyCode
    gIndice = cTipo.ListIndex
End Sub

Private Sub cTipo_KeyPress(KeyAscii As Integer)
    cTipo.ListIndex = gIndice
    
    If KeyAscii = vbKeyReturn And cTipo.ListIndex <> -1 Then
        If cTipo.ItemData(cTipo.ListIndex) = TipoDocumento.Remito Then
            CamposRemito False, Inactivo
            Foco tNumero
        Else
            CamposRemito True, Blanco
            Foco cSucursal
        End If
    End If
    
End Sub

Private Sub cTipo_KeyUp(KeyCode As Integer, Shift As Integer)
    ComboKeyUp cTipo
End Sub

Private Sub cTipo_LostFocus()
    
    gIndice = -1
    cTipo.SelLength = 0
    
    'If cTipo.ListIndex <> -1 Then
    '    If cTipo.ItemData(cTipo.ListIndex) = TipoDocumento.Remito Then CamposRemito False, Inactivo Else CamposRemito True, Blanco
    'End If
    
End Sub

Private Sub CamposRemito(Activo As Boolean, Color As Variant)
    cSucursal.Enabled = Activo
    cSucursal.BackColor = Color
    tSerie.Enabled = Activo
    tSerie.BackColor = Color
    cSucursal.Text = ""
    tSerie.Text = ""
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()

     SetearLView lvValores.Grilla Or lvValores.FullRow, lArticulo
     
    'Cargo combo con tipos de docuemento--------------------------------------
    cTipo.AddItem Trim(DocContado)
    cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.Contado
    cTipo.AddItem Trim(DocCredito)
    cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.Credito
    cTipo.AddItem Trim(DocRemito)
    cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.Remito
    
    'Cargo Sucursales---------------------------------------------------------------------------
    Cons = "Select SucCodigo, SucAbreviacion from Sucursal Order by SucAbreviacion"
    CargoCombo Cons, cSucursal, ""
    '-----------------------------------------------------------------------------------------------
    
    Fletes = CargoArticulosDeFlete
    EstadoEntregando False
    
    ReDim arrNroSerie(0)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    'Forms(Forms.Count - 2).SetFocus
    InicioEMercaderia.SetFocus
    
End Sub

Private Sub Label1_Click()
    Foco tCBarra
End Sub

Private Sub Label2_Click()
    Foco tArticulo
End Sub

Private Sub Label4_Click()
    Foco cSucursal
End Sub

Private Sub Label5_Click()
    Foco tUsuario
End Sub

Private Sub Label6_Click()
    If tSerie.Enabled Then Foco tSerie Else Foco tNumero
End Sub

Private Sub Label7_Click()
    Foco tComentario
End Sub

Private Sub Label8_Click()
    Foco cTipo
End Sub

Private Sub lArticulo_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeySubtract
            If lArticulo.ListItems.Count = 0 Or lArticulo.SelectedItem.Index = -1 Then Exit Sub
            Set itmX = lArticulo.SelectedItem
            If CCur(itmX.SubItems(3)) > 0 Then
            
                If Mid(itmX.Key, 1, 1) = "S" Then  'Nro de Serie-------------------------------------!!!!!
                    Dim aNroSerie As String: aNroSerie = ""
                    Do While aNroSerie = ""
                        aNroSerie = InputBox("Ingrese el número de nerie del artículo devuelto.", "(" & Trim(itmX.Text) & ") " & Trim(itmX.SubItems(1)))
                        If Trim(aNroSerie) <> "" Then
                            If Not arrEliminoElemento(itmX.Tag, Trim(aNroSerie)) Then aNroSerie = ""
                        End If
                    Loop
                End If      '-------------------------------------------------------------------------------------------------------
                
                itmX.SubItems(2) = CCur(itmX.SubItems(2)) + 1
                itmX.SubItems(3) = CCur(itmX.SubItems(3)) - 1
            End If
            
    End Select
    
End Sub

Private Sub tArticulo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        On Error GoTo errEntregar
        
        'Si el Codigo ingresado es <= 7 es de CGSA sino es el Codigo de Barras del Artículo
        'Si tiene ingresada la barra de cantidad (/) ---> tambien es de CGSA
        
        If InStr(tArticulo.Text, "/") = 0 Then
            If Len(tArticulo.Text) <= 7 Then
                'Codigo de CGSA
                If IsNumeric(tArticulo.Text) Then EntregoArticulo Trim(tArticulo.Text), 1, False Else: MsgBox "El código ingresado no es correcto.", vbExclamation, "ATENCIÓN"
            Else
                'BarCode
                EntregoArticulo Trim(tArticulo.Text), 1, True
            End If
        
        Else
            'Codigo de CGSA
            If IsNumeric(Mid(tArticulo.Text, 1, InStr(tArticulo.Text, "/") - 1)) And IsNumeric(Mid(tArticulo.Text, InStr(tArticulo.Text, "/") + 1, Len(tArticulo.Text))) Then
                EntregoArticulo Trim(Mid(tArticulo.Text, 1, InStr(tArticulo.Text, "/") - 1)), CCur(Mid(tArticulo.Text, InStr(tArticulo.Text, "/") + 1, Len(tArticulo.Text))), False
            End If
        End If
        
    End If
    Exit Sub

errEntregar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al interpretar al código."
End Sub

Private Sub tCBarra_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tCBarra.Text) = "" Then
            cTipo.SetFocus
        Else
            FormatoBarras Trim(tCBarra.Text)
            On Error Resume Next
            If tCBarra.Enabled Then Foco tCBarra
        End If
    End If
    
End Sub

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub tNumero_GotFocus()
    tNumero.SelStart = 0
    tNumero.SelLength = Len(tNumero.Text)
End Sub

Private Sub tNumero_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        'Valido los datos ingresados para búsqueda manual-----------------------------------------------------
        If cTipo.ListIndex = -1 Then
            MsgBox "Debe seleccionar el tipo de documento a buscar.", vbExclamation, "ATENCIÓN"
            Foco cTipo: Exit Sub
        End If
        
        If cTipo.ItemData(cTipo.ListIndex) <> TipoDocumento.Remito Then
            If cSucursal.ListIndex = -1 Or Trim(tSerie.Text) = "" Or IsNumeric(tSerie.Text) Then
                MsgBox "Los datos ingresados no son correctos. Verifique.", vbExclamation, "ATENCIÓN"
                Foco cSucursal: Exit Sub
            End If
            
            If Not IsNumeric(tNumero.Text) Then
                MsgBox "Debe ingresar el número del documento.", vbExclamation, "ATENCIÓN"
                Foco cSucursal: Exit Sub
            End If
        End If
        '--------------------------------------------------------------------------------------------------------------
        Screen.MousePointer = 11
        On Error GoTo errCargar
        
        gTipo = CInt(cTipo.ItemData(cTipo.ListIndex))
        
        If gTipo = TipoDocumento.Remito Then
            BuscoRemito CLng(tNumero.Text)
        Else
            BuscoDocumento gTipo, cSucursal.ItemData(cSucursal.ListIndex), tSerie.Text, CLng(tNumero.Text)
        End If
        
        Screen.MousePointer = 0
        
    End If
    Exit Sub

errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del documento."
End Sub

Private Sub BuscoDocumento(Optional Tipo As Integer, Optional Sucursal As Long, Optional Serie As String, Optional Numero As Long, Optional Codigo As Long = 0)

    If Codigo <> 0 Then
        Cons = "Select * from Documento Where DocCodigo = " & Codigo & " And DocTipo = " & Tipo
    Else
        Cons = "Select * from Documento" _
                & " Where DocSucursal = " & Sucursal _
                & " And DocTipo = " & Tipo _
                & " And DocSerie = '" & Trim(Serie) & "'" _
                & " And DocNumero = " & Numero
    End If
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not RsAux.EOF Then
        gDocumento = RsAux!DocCodigo
        gFechaDocumento = RsAux!DocFModificacion
        gFechaEmision = RsAux!DocFecha
        
        'Cargo los datos del Documento Seleccionado----------------
        BuscoCodigoEnCombo cTipo, RsAux!DocTipo
        BuscoCodigoEnCombo cSucursal, RsAux!DocSucursal
        tSerie.Text = Trim(RsAux!DocSerie)
        tNumero.Text = RsAux!DocNumero
        '-------------------------------------------------------------------
        
        CargoCliente RsAux!DocCliente
        
        Cons = "Select ArtID, ArtCodigo, ArtBarCode, ArtNombre, RenCantidad Cantidad, RenARetirar ARetirar, ArtNroSerie, ArtTipo from Renglon, Articulo" _
                & " Where RenDocumento = " & RsAux!DocCodigo _
                & " And RenArticulo = ArtID "
            
        CargoArticulos Cons, False
        
        If RsAux!DocAnulado Then
            Screen.MousePointer = 0
            MsgBox "El documento ingresado ha sido anulado. Verifique", vbCritical, "DOCUMENTO ANULADO"
            EstadoEntregando False
        Else
            If Not IsNull(RsAux!DocPendiente) Then
                Screen.MousePointer = 0
                MsgBox "La mercadería está pendiente de entrega. Verifique", vbInformation, "ATENCIÓN"
                EstadoEntregando False
            Else
                Foco tUsuario
            End If
        End If
    Else
        Screen.MousePointer = 0
        gDocumento = 0
        MsgBox "No existe un documento para las características ingresadas.", vbExclamation, "ATENCIÓN"
    End If
    
    RsAux.Close
    
End Sub

Private Sub BuscoRemito(Numero As Long)

    Cons = "Select * from Remito, Documento" _
            & " Where RemCodigo = " & Numero _
            & " And RemDocumento = DocCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)


    If Not RsAux.EOF Then
        gDocumento = RsAux!DocCodigo
        gFechaDocumento = RsAux!DocFModificacion     'Siempre guardo la del Documento
        gFechaEmision = RsAux!DocFecha
        
        'Cargo los datos del Documento Seleccionado----------------
        BuscoCodigoEnCombo cTipo, TipoDocumento.Remito
        cSucursal.Text = ""
        tSerie.Text = ""
        tNumero.Text = RsAux!RemCodigo
        '-------------------------------------------------------------------
        
        CargoCliente RsAux!DocCliente
        
        Cons = "Select ArtID, ArtCodigo, ArtBarCode, ArtNombre, RReCantidad Cantidad, RReAEntregar ARetirar, ArtNroSerie, ArtTipo " _
                & " From RenglonRemito, Articulo" _
                & " Where RReRemito = " & RsAux!RemCodigo _
                & " And RReArticulo = ArtID "
            
        CargoArticulos Cons, True
        
        If RsAux!DocAnulado Then
            Screen.MousePointer = 0
            MsgBox "El documento ingresado ha sido anulado. Verifique", vbCritical, "DOCUMENTO ANULADO"
            EstadoEntregando False
        Else
            If Not IsNull(RsAux!DocPendiente) Then
                Screen.MousePointer = 0
                MsgBox "La mercadería está pendiente de entrega. Verifique", vbInformation, "ATENCIÓN"
                EstadoEntregando False
            Else
                Foco tUsuario
            End If
        End If
        
    Else
        Screen.MousePointer = 0
        gDocumento = 0
        MsgBox "No existe un remito para las características ingresadas.", vbExclamation, "ATENCIÓN"
    End If
    
    RsAux.Close

End Sub

Private Sub CargoCliente(Cliente As Long)
    
    On Error GoTo errCliente
     Cons = "Select CliCiRuc, CliTipo, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2) " _
           & " From Cliente, CPersona " _
           & " Where CliCodigo = " & Cliente _
           & " And CliCodigo = CPeCliente " _
                                                & " UNION " _
           & " Select CliCiRuc, CliTipo, Nombre = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')')  " _
           & " From Cliente, CEmpresa " _
           & " Where CliCodigo = " & Cliente _
           & " And CliCodigo = CEmCliente"

    Set rsXX = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    lCiRuc.Caption = "S/D"
    If Not rsXX.EOF Then
        If rsXX!CliTipo = TipoCliente.Persona Then
            If Not IsNull(rsXX!CliCIRuc) Then lCiRuc.Caption = clsGeneral.RetornoFormatoCedula(rsXX!CliCIRuc)
        Else
            If Not IsNull(rsXX!CliCIRuc) Then lCiRuc.Caption = clsGeneral.RetornoFormatoRuc(Trim(rsXX!CliCIRuc))
        End If
    End If
    lTitular.Caption = Trim(rsXX!Nombre)
    rsXX.Close
    Exit Sub
    
errCliente:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del cliente."
End Sub

Private Sub tSerie_GotFocus()
    tSerie.SelStart = 0
    tSerie.SelLength = Len(tSerie.Text)
End Sub

Private Sub tSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn And Not IsNumeric(tSerie.Text) And Trim(tSerie.Text) <> "" Then Foco tNumero
End Sub

'--------------------------------------------------------------------------------------------------------
'   Los campos seleccionados en la cosnulta se deben renombrar como:
'       ArtID, ArtCodigo, ArtBarCode, ArtNombre, ARetirar
'--------------------------------------------------------------------------------------------------------
Private Sub CargoArticulos(Consulta As String, EsRemito As Boolean)

Dim aADevolver As Integer
Dim EnRemito As Currency, EnNota As Currency, EnEnvio As Currency, EnDevolucion As Currency

    On Error GoTo errCargar
    lArticulo.ListItems.Clear
    aADevolver = 0
         
    Set rsXX = cBase.OpenResultset(Consulta, rdOpenDynamic, rdConcurValues)
    Do While Not rsXX.EOF
        'Si no es flete lo cargo
        If InStr(Fletes, rsXX!ArtID & ",") = 0 Then
            
            Set itmX = lArticulo.ListItems.Add(Text:=Trim(rsXX!ArtCodigo))
            itmX.Tag = rsXX!ArtID
            
            itmX.SubItems(1) = Trim(rsXX!ArtNombre)
            
            If Not EsRemito Then
                EnRemito = ArticulosEnRemito(gDocumento, rsXX!ArtID)
                EnEnvio = ArticulosEnEnvio(gDocumento, rsXX!ArtID)
                EnNota = ArticulosEnNota(gDocumento, rsXX!ArtID)
                EnDevolucion = ArticulosEnDevolucion(gDocumento, rsXX!ArtID)
            Else
                EnRemito = 0
                EnNota = 0
                EnEnvio = 0
                EnDevolucion = 0
            End If
        
            itmX.SubItems(2) = rsXX!Cantidad - rsXX!ARetirar - EnRemito - EnNota - EnEnvio - EnDevolucion
            aADevolver = aADevolver + CCur(itmX.SubItems(2))
            
            itmX.SubItems(3) = "0"
            
            If CCur(itmX.SubItems(2)) <> 0 Then itmX.SmallIcon = "si" Else: itmX.SmallIcon = "no"
            
            If Not IsNull(rsXX!ArtBarCode) Then itmX.SubItems(4) = Trim(rsXX!ArtBarCode)
            
            itmX.SubItems(5) = rsXX!ArtTipo
            
            'Si el Art requiere NroSerie el Key se arma con "S + id_Articulo", sino con "N + id_Articulo"
            If rsXX!ArtNroSerie Then itmX.Key = "S" & rsXX!ArtID Else itmX.Key = "N" & rsXX!ArtID
        End If
        
        rsXX.MoveNext
    Loop
    rsXX.Close
    
    lAEntregar.Caption = "Total de Artículos a Devolver:      " & aADevolver & " "
    
    If aADevolver = 0 Then
        MsgBox "El documento no tiene artículos para devolver." & vbCrLf & "Están asociados a otro documento o no se han entregado.", vbInformation, "ATENCIÓN"
        EstadoEntregando False
    Else
        'Valido que no sea un documento de Servicios---------------------------------------------
        Dim rsSer As rdoResultset
        Cons = "Select * From Servicio Where SerDocumento = " & gDocumento
        Set rsSer = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsSer.EOF Then
            MsgBox "Este documento está asociado a el servicio Nº " & rsSer!SerCodigo & Chr(vbKeyReturn) & "No se podrán realizar devoluciones de artículos.", vbInformation, "Factura de Servicios."
            EstadoEntregando False
        Else
            EstadoEntregando True
        End If
        rsSer.Close
        '-------------------------------------------------------------------------------------------------
        
    End If
    Exit Sub

errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los artículos del documento."
End Sub


Private Function ArticulosEnRemito(Documento As Long, Articulo As Long) As Currency
    
Dim rs1 As rdoResultset

    ArticulosEnRemito = 0
    On Error GoTo errArt
    Cons = "Select Sum(RReCantidad) from Remito, RenglonRemito" _
           & " Where RemDocumento = " & gDocumento _
           & " And RemCodigo = RReRemito" _
           & " And RReArticulo = " & Articulo
    Set rs1 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rs1.EOF Then
        If Not IsNull(rs1(0)) Then ArticulosEnRemito = rs1(0)
    End If
    rs1.Close
    Exit Function
    
errArt:
End Function

Private Function ArticulosEnEnvio(Documento As Long, Articulo As Long) As Currency
    
Dim rs1 As rdoResultset

    ArticulosEnEnvio = 0
    On Error GoTo errArt
    Cons = "Select Sum(REvCantidad) from Envio, RenglonEnvio" _
           & " Where EnvDocumento = " & gDocumento _
           & " And EnvCodigo = REvEnvio" _
           & " And REvArticulo = " & Articulo
    Set rs1 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rs1.EOF Then
        If Not IsNull(rs1(0)) Then ArticulosEnEnvio = rs1(0)
    End If
    rs1.Close
    Exit Function
    
errArt:
End Function

Private Function ArticulosEnNota(Documento As Long, Articulo As Long) As Currency
    
Dim rs1 As rdoResultset

    ArticulosEnNota = 0
    On Error GoTo errArt
    
    Cons = "Select Sum(RenCantidad) from Nota, Renglon, Documento" _
           & " Where NotFactura = " & gDocumento _
           & " And NotNota = RenDocumento" _
           & " And RenArticulo = " & Articulo _
           & " And NotNota = DocCodigo And DocAnulado = 0"
    Set rs1 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rs1.EOF Then
        If Not IsNull(rs1(0)) Then ArticulosEnNota = rs1(0)
    End If
    rs1.Close
    Exit Function
    
errArt:
End Function

Private Function ArticulosEnDevolucion(Documento As Long, lnArticulo As Long) As Integer
    On Error GoTo ErrCAER
    Dim rs1 As rdoResultset
    
    ArticulosEnDevolucion = 0
    
    Cons = "Select * From Devolucion  " & _
               " Where DevNota Is Null " & _
               " And DevArticulo = " & lnArticulo & _
               " And DevFactura = " & Documento & _
               " And DevAnulada Is Null"
    
    Set rs1 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not rs1.EOF Then ArticulosEnDevolucion = rs1!DevCantidad
    rs1.Close
    
    Exit Function
    
ErrCAER:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al buscar la cantidad de artículos en devoluciones.", Err.Description
End Function


Private Sub EstadoEntregando(Estado As Boolean)

    tCBarra.Enabled = Not Estado
    cTipo.Enabled = Not Estado
    cSucursal.Enabled = Not Estado
    tSerie.Enabled = Not Estado
    tNumero.Enabled = Not Estado
    
    tArticulo.Enabled = Estado
    tUsuario.Enabled = Estado
    tComentario.Enabled = Estado
    
    If Estado Then
        tArticulo.BackColor = Blanco
        tUsuario.BackColor = Blanco
        tComentario.BackColor = Blanco
        
        tCBarra.BackColor = Inactivo
        cTipo.BackColor = Inactivo
        cSucursal.BackColor = Inactivo
        tSerie.BackColor = Inactivo
        tNumero.BackColor = Inactivo
    Else
        tArticulo.BackColor = Inactivo
        tUsuario.BackColor = Inactivo
        tComentario.BackColor = Inactivo
        
        tCBarra.BackColor = Blanco
        cTipo.BackColor = Blanco
        cSucursal.BackColor = Blanco
        tSerie.BackColor = Blanco
        tNumero.BackColor = Blanco
    End If
    
End Sub

Private Sub EntregoArticulo(Codigo As String, Cantidad As Currency, Optional EsBarCode As Boolean = True)

Dim sEntrego As Boolean
Dim aCodigo As String

    sEntrego = False
    'Bajo el Articulo de la lista
    For Each itmX In lArticulo.ListItems
        If EsBarCode Then aCodigo = Trim(itmX.SubItems(4)) Else: aCodigo = Trim(itmX.Text)
        
        If Trim(aCodigo) = Trim(Codigo) Then
        
            If CCur(itmX.SubItems(2)) - Cantidad >= 0 Then
                'Entrego el Articulo-------------------------------------------------
                itmX.SubItems(2) = CCur(itmX.SubItems(2)) - Cantidad
                itmX.SubItems(3) = CCur(itmX.SubItems(3)) + Cantidad
                If CCur(itmX.SubItems(2)) = 0 Then itmX.SmallIcon = "no"
                
                If Cantidad = 1 And Mid(itmX.Key, 1, 1) = "S" Then 'Nro de Serie-------------------------------------!!!!!
                    Dim aNroSerie As String: aNroSerie = ""
                    Do While aNroSerie = ""
                        aNroSerie = InputBox("Ingrese el número de nerie del artículo devuelto.", "(" & Trim(itmX.Text) & ") " & Trim(itmX.SubItems(1)))
                        If Trim(aNroSerie) <> "" Then
                            If Not arrAgregoElemento(itmX.Tag, Trim(aNroSerie)) Then aNroSerie = ""
                        End If
                    Loop
                End If
                '-------------------------------------------------------------------------------------------------------------------
                
            Else
                'No hay para entregar----------------------------------------------
                MsgBox "El / los " & Trim(itmX.SubItems(1)) & " ya se han devueltos.", vbExclamation, "ATENCIÓN"
            End If
            
            sEntrego = True
            Exit For
        End If
    Next
    
    tArticulo.Text = ""
    
    If Not sEntrego Then
        MsgBox "El artículo ingresado no figura en la lista. No se puede devolver con éste documento.", vbCritical, "ATENCIÓN"
    End If
       
    'Si ya se entregaron TODOS Grabo------------------------------------------------
    sEntrego = True
    For Each itmX In lArticulo.ListItems
        If itmX.SmallIcon = "si" Then sEntrego = False: Exit For
    Next
    If sEntrego Then Foco tComentario
    '----------------------------------------------------------------------------------------
    
End Sub

Private Sub AccionGrabar()

    'Valido si hay articulos devueltos------------------------------------
    Dim sHay As Boolean
    sHay = False
    For Each itmX In lArticulo.ListItems
        If itmX.SubItems(3) <> 0 Then sHay = True: Exit For
    Next
    '--------------------------------------------------------------------------
    
    If Not sHay Then        'Accion Cancelar----------
        If MsgBox("No hay artículos devueltos. Cancela el ingreso.", vbQuestion + vbYesNo) = vbYes Then
            LimpioParaNuevaEntrega
        Else
            lArticulo.SetFocus
        End If
        Exit Sub
    End If
    
    If tUsuario.Tag = "0" Or tUsuario.Tag = "" Then
        MsgBox "Ingrese al dígito del usuario que recibe la mercadería.", vbExclamation, "ATENCIÓN"
        Foco tUsuario
        Exit Sub
    End If
    
    If MsgBox("Confirma almacenar la devolución de entrega.", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    
    Dim gSucesoUsr As Long, gSucesoDef As String
    gSucesoUsr = 0
    'Suceso si el documento no es del dia.-----------------------------------------------------------------------------------
    If Format(gFechaEmision, "dd/mm/yyyy") <> Format(gFechaServidor, "dd/mm/yyyy") Then
        Dim objSuceso As New clsSuceso
        objSuceso.ActivoFormulario CLng(tUsuario.Tag), "Devolución de Mercadería", cBase
        gSucesoUsr = objSuceso.RetornoValor(Usuario:=True)
        gSucesoDef = objSuceso.RetornoValor(Defensa:=True)
        Set objSuceso = Nothing
        Me.Refresh
        
        If gSucesoUsr = 0 Then Exit Sub
    End If
    
    '------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errorBT
    FechaDelServidor    'Saco la fecha del Servidor
    Screen.MousePointer = 11

    aTexto = 0
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
    On Error GoTo errorET
    
    'Antes veo si la fecha es la misma-------------------------------------------------------------------------------------
    Cons = "Select * from Documento Where DocCodigo = " & gDocumento
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux!DocFModificacion <> gFechaDocumento Then
        aTexto = "El documento seleccionado ha sido modificado por otra terminal. Vuelva a consultar."
        RsAux.Close
        GoTo errorET
        Exit Sub
    Else
        RsAux.Edit
        RsAux!DocFModificacion = Format(gFechaServidor, FormatoFH)
        RsAux.Update
    End If
    RsAux.Close '-----------------------------------------------------------------------------------------------------------
    
    GraboDatosTablas
    GraboProductosVendidos gDocumento
    
    If gSucesoUsr <> 0 Then
        clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.VariosStock, paCodigoDeTerminal, gSucesoUsr, gDocumento, Descripcion:="Devolución - El doc. no es del día.", Defensa:=gSucesoDef
    End If
    
    cBase.CommitTrans    'Fin de la TRANSACCION------------------------------------------
    Screen.MousePointer = 0
    
    LimpioParaNuevaEntrega
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
    If aTexto = "" Then aTexto = "No se ha podido realizar la transacción. Reintente la operación."
    clsGeneral.OcurrioError aTexto
    Exit Sub
End Sub


Private Sub GraboDatosTablas()

Dim aDocumento As Long

    If gTipo = TipoDocumento.Remito Then aDocumento = CLng(tNumero.Text) Else aDocumento = gDocumento
    
    For Each itmX In lArticulo.ListItems
    
        If CCur(itmX.SubItems(3)) > 0 Then
            'Actualizo los datos en tabla Renglones------------------------------------------------------------------------------------
            If gTipo = TipoDocumento.Remito Then
                Cons = "Update RenglonRemito " _
                    & " Set RReAEntregar = RReAEntregar + " & CCur(itmX.SubItems(3)) _
                    & " Where RReRemito = " & aDocumento _
                    & " And RReArticulo = " & itmX.Tag
            Else
                Cons = "Update Renglon " _
                    & " Set RenARetirar = RenARetirar + " & CCur(itmX.SubItems(3)) _
                    & " Where RenDocumento = " & aDocumento _
                    & " And RenArticulo = " & itmX.Tag
            End If
            cBase.Execute Cons
            '-------------------------------------------------------------------------------------------------------------------------------
            If Val(itmX.SubItems(5)) <> paTipoArticuloServicio Then
                'Marco el Alta del STOCK AL LOCAL
                'Genero Movimiento
                MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, paCodigoDeSucursal, CLng(itmX.Tag), CCur(itmX.SubItems(3)), paEstadoArticuloEntrega, 1, gTipo, aDocumento
                'Bajo del Stock en Local
                MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, CLng(itmX.Tag), CCur(itmX.SubItems(3)), paEstadoArticuloEntrega, 1
                
                'Marco el Movimiento del STOCK VIRTUAL
                'Genero Movimiento
                MarcoMovimientoStockEstado CLng(tUsuario.Tag), CLng(itmX.Tag), CCur(itmX.SubItems(3)), TipoMovimientoEstado.ARetirar, 1, gTipo, aDocumento, paCodigoDeSucursal
                'Bajo del Stock Total
                MarcoMovimientoStockTotal CLng(itmX.Tag), TipoEstadoMercaderia.Virtual, TipoMovimientoEstado.ARetirar, CCur(itmX.SubItems(3)), 1
            End If
        End If
    Next
    
End Sub

Private Sub GraboProductosVendidos(idDocumento As Long)

    Dim rsPV As rdoResultset
    
    For I = LBound(arrNroSerie) To UBound(arrNroSerie)
        If arrNroSerie(I).Articulo <> -1 And Trim(arrNroSerie(I).NroSerie) <> "" Then
            
            Cons = "Select * from ProductosVendidos Where PVeDocumento = " & idDocumento & _
                        " And PVeArticulo = " & arrNroSerie(I).Articulo & _
                        " And PVeNSerie = '" & Trim(arrNroSerie(I).NroSerie) & "'"
            Set rsPV = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rsPV.EOF Then rsPV.Delete
            rsPV.Close
    
        End If
    Next
    
End Sub

Private Sub LimpioParaNuevaEntrega()

    EstadoEntregando False
    lAEntregar.Caption = ""
    lArticulo.ListItems.Clear
    tCBarra.Text = ""
    cTipo.Text = ""
    cSucursal.Text = ""
    tSerie.Text = ""
    tNumero.Text = ""
    
    tUsuario.Text = ""
    tUsuario.Tag = 0
    tComentario.Text = ""
    
    lCiRuc.Caption = "N/D"
    lTitular.Caption = "N/D"
    
    Foco tCBarra
    gDocumento = 0
    
    ReDim arrNroSerie(0)
        
End Sub

'----------------------------------------------------------------------------------
'   Interpreta el Texto del Codigo de Barras
'   Formato:    XDXXXX          TipoDocumento   D(Separador)     Numero de Documento
'----------------------------------------------------------------------------------
Private Sub FormatoBarras(Texto As String)

Dim aCodDoc As Long
    
    On Error GoTo errInt
    Texto = UCase(Texto)
    gTipo = CLng(Mid(Texto, 1, InStr(Texto, "D") - 1))
    aCodDoc = CLng(Trim(Mid(Texto, InStr(Texto, "D") + 1, Len(Texto))))
    
    Select Case gTipo
        Case TipoDocumento.Remito:  BuscoRemito CLng(tNumero.Text)
        
        Case TipoDocumento.Contado, TipoDocumento.Credito: BuscoDocumento Tipo:=gTipo, Codigo:=aCodDoc
        
        Case Else:  MsgBox "El código de barras ingresado no es correcto. El documento no coincide con los predefinidos.", vbCritical, "ATENCIÓN"
    End Select
        
    Exit Sub
    
errInt:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al interpretar el código de barras."
End Sub

Private Sub tUsuario_Change()
    tUsuario.Tag = 0
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tUsuario.Text) Then
            tUsuario.Tag = BuscoUsuario(tUsuario.Text)
            If tUsuario.Tag <> 0 And tUsuario.Text <> "" Then Foco tArticulo
        End If
    End If
    
End Sub

Private Function arrAgregoElemento(aIdArticulo As Long, aSerie As String) As Boolean
    
    On Error GoTo errAgregar
    arrAgregoElemento = False
    If arrBuscoElemento(aIdArticulo, aSerie) <> 0 Then
        MsgBox "El nro. de serie ingresado ya fue devuelto !!!!." & vbCrLf & vbCrLf & "Nº Serie: " & Trim(aSerie), vbExclamation, "Artículo Devuelto"
        Exit Function
    End If

    Dim aIdxC As Integer
    aIdxC = UBound(arrNroSerie) + 1
    ReDim Preserve arrNroSerie(aIdxC)
        
    arrNroSerie(aIdxC).Articulo = aIdArticulo
    arrNroSerie(aIdxC).NroSerie = Trim(aSerie)
    
    arrAgregoElemento = True

errAgregar:
End Function

Private Function arrEliminoElemento(aIdArticulo As Long, aSerie As String) As Boolean

    arrEliminoElemento = False
    If arrBuscoElemento(aIdArticulo, aSerie) = 0 Then
        MsgBox "Este artículo no fue devuelto !!!!." & vbCrLf & vbCrLf & "Nº Serie: " & Trim(aSerie), vbExclamation, "Artículo NO Devuelto"
        Exit Function
    End If

    Dim I As Integer
    For I = LBound(arrNroSerie) To UBound(arrNroSerie)
        If aIdArticulo = arrNroSerie(I).Articulo And UCase(aSerie) = UCase(arrNroSerie(I).NroSerie) Then
            arrNroSerie(I).Articulo = -1
            arrNroSerie(I).NroSerie = ""
            Exit For
        End If
    Next
    
    arrEliminoElemento = True
    
End Function

Private Function arrBuscoElemento(aIdArticulo As Long, aSerie As String) As Long
    
    On Error GoTo errB
    arrBuscoElemento = 0
    Dim I As Integer
    For I = LBound(arrNroSerie) To UBound(arrNroSerie)
        If aIdArticulo = arrNroSerie(I).Articulo And UCase(aSerie) = UCase(arrNroSerie(I).NroSerie) Then
            arrBuscoElemento = I: Exit Function
        End If
    Next
errB:
End Function


