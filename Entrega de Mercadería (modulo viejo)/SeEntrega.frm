VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form SeEntrega 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrega de Mercadería"
   ClientHeight    =   4155
   ClientLeft      =   3555
   ClientTop       =   2970
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
   Icon            =   "SeEntrega.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   6840
   Begin VB.CommandButton bCancelar 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   5880
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton bDFactura 
      Caption         =   "&Detalle de Factura"
      Height          =   315
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton bFinalizar 
      Caption         =   "Finali&zar"
      Height          =   315
      Left            =   4920
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox tArticulo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3000
      MaxLength       =   30
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
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
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   3201
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
         Text            =   "A Entregar"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Entregados"
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
   Begin VB.Label lParcial 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ENTREGA PARCIAL !!! "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   840
      Width           =   4575
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5940
      Top             =   -300
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
            Picture         =   "SeEntrega.frx":0442
            Key             =   "si"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SeEntrega.frx":075C
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
      TabIndex        =   13
      Top             =   3840
      Width           =   6600
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   1320
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
      TabIndex        =   11
      Top             =   840
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
      TabIndex        =   10
      Top             =   1320
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
      TabIndex        =   9
      Top             =   1080
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "CI/RUC:"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&ENTREGANDO:"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   1725
      Width           =   1215
   End
   Begin VB.Label lDocumento 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "CONTADO A-985010"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2580
      TabIndex        =   2
      Top             =   315
      Width           =   4155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      Height          =   900
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   735
      Width           =   6615
   End
End
Attribute VB_Name = "SeEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum EstadoS            'Enum de Servicios---------------------------
    Anulado = 0
    Visita = 1
    Retiro = 2
    Taller = 3
    Entrega = 4
    Cumplido = 5
End Enum                  '------------------------------------------------------

Dim gDocumento As Long
Dim gRemito As Long
Dim gTipo As Integer
Dim gFechaDocumento As Date

Dim rsXX As rdoResultset
Dim Fletes As String

Dim gTeclaFuncion As Long
Dim gCumplirServicio As Long

'------------------------------------------------------------------
Private Type typNroSerie
    Articulo As Long
    NroSerie As String
End Type

Dim arrNroSerie() As typNroSerie
'------------------------------------------------------------------

Public Property Get pTecla() As Long
    pTecla = gTeclaFuncion
End Property
Public Property Let pTecla(Codigo As Long)
    gTeclaFuncion = Codigo
End Property

Private Sub bCancelar_Click()
    
    If gDocumento <> 0 Then
        
        'Valido si hay articulos entregados------------------------------------
        Dim sHay As Boolean, aDefensa As String
        sHay = False: aDefensa = ""
        For Each itmX In lArticulo.ListItems
            If itmX.SubItems(3) <> 0 Then
                sHay = True
                aDefensa = aDefensa & Trim(itmX.SubItems(3)) & " " & Trim(itmX.SubItems(1)) & "; "
            End If
        Next
        If Len(Trim(aDefensa)) > 2 Then
            aDefensa = "Pasó: " & Mid(aDefensa, 1, Len(aDefensa) - 2)
        Else
            aDefensa = "No se pasaron artículos."
        End If
        '--------------------------------------------------------------------------
        If sHay Then
            If MsgBox("Está seguro que desea cancelar.", vbQuestion + vbYesNo + vbDefaultButton2, "Cancelar Entrega") = vbNo Then Exit Sub
        End If
        
        If Val(lAEntregar.Tag) <> 0 Then
            On Error Resume Next
            Dim aTxt As String
            aTxt = "Cancela Entrega de Merc. "
            aTxt = aTxt & Trim(lDocumento.Caption)
            
            Screen.MousePointer = 11
            'Registro Suceso
            FechaDelServidor
            clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.VariosStock, paCodigoDeTerminal, Val(Me.Tag), gDocumento, , aTxt, aDefensa
        End If
        
        Screen.MousePointer = 0
    End If
    
    LimpioParaNuevaEntrega False
End Sub

Private Sub bDFactura_Click()
    
    If gDocumento = 0 Then Exit Sub
    EjecutarApp pathApp & "\Detalle de factura", CStr(gDocumento)
    
End Sub

Private Sub bFinalizar_Click()
    AccionGrabar
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
    If KeyCode = vbKeyEscape Then Me.WindowState = vbMinimized
        
End Sub

Private Sub Form_Load()

    lDocumento.Caption = ""
    SetearLView lvValores.Grilla Or lvValores.FullRow, lArticulo
    
    Fletes = CargoArticulosDeFlete
    EstadoEntregando False
    lParcial.Visible = False
    
    ReDim arrNroSerie(0)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If gDocumento <> 0 Then
        MsgBox "Hay un documento pendiente de entrega. Termine la tarea antes de salir.", vbInformation, "ATENCIÓN"
        Cancel = 1
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    Select Case gTeclaFuncion
        Case vbKeyF1: InicioEMercaderia.MnuF1.Caption = "Vacío........"
        Case vbKeyF2: InicioEMercaderia.MnuF2.Caption = "Vacío........"
        Case vbKeyF3: InicioEMercaderia.MnuF3.Caption = "Vacío........"
        Case vbKeyF4: InicioEMercaderia.MnuF4.Caption = "Vacío........"
        Case vbKeyF5: InicioEMercaderia.MnuF5.Caption = "Vacío........"
        Case vbKeyF6: InicioEMercaderia.MnuF6.Caption = "Vacío........"
        Case vbKeyF7: InicioEMercaderia.MnuF7.Caption = "Vacío........"
        Case vbKeyF8: InicioEMercaderia.MnuF8.Caption = "Vacío........"
        Case vbKeyF9: InicioEMercaderia.MnuF9.Caption = "Vacío........"
        Case vbKeyF11: InicioEMercaderia.MnuF11.Caption = "Vacío........"
        Case vbKeyF12: InicioEMercaderia.MnuF12.Caption = "Vacío........"
    End Select
    InicioEMercaderia.SetFocus
    
End Sub

Private Sub Label1_Click()
    Foco tCBarra
End Sub

Private Sub Label2_Click()
    Foco tArticulo
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
                itmX.SmallIcon = "si"
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
        If Trim(tCBarra.Text) <> "" Then
            FormatoBarras Trim(tCBarra.Text)
            On Error Resume Next
            If tCBarra.Enabled Then Foco tCBarra
        End If
    End If
    
End Sub


Private Sub BuscoDocumento(Optional Tipo As Integer, Optional Sucursal As Long, Optional Serie As String, Optional Numero As Long, Optional Codigo As Long = 0)

Dim idCliente As Long

    Screen.MousePointer = 11
    gCumplirServicio = 0
    idCliente = 0
    
    lDocumento.Caption = ""
    
    If Codigo <> 0 Then
        Cons = "Select * from Documento Where DocCodigo = " & Codigo '& " And DocTipo = " & Tipo
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
        
        lDocumento.Caption = UCase(NombreDocumento(RsAux!DocTipo) & " " & Trim(RsAux!DocSerie) & "-" & RsAux!DocNumero)
                
        idCliente = RsAux!DocCliente
        CargoCliente idCliente
        
        Cons = "Select ArtID, ArtCodigo, ArtBarCode, ArtNombre, RenARetirar ARetirar, ArtNroSerie, ArtTipo from Renglon, Articulo" _
                & " Where RenDocumento = " & RsAux!DocCodigo _
                & " And RenArticulo = ArtID "
            
        CargoArticulos Cons
        
        If RsAux!DocAnulado Then
            gCumplirServicio = 0
            Screen.MousePointer = 0
            MsgBox "El documento ingresado ha sido anulado. Verifique", vbCritical, "DOCUMENTO ANULADO"
            EstadoEntregando False
        Else
            If Not IsNull(RsAux!DocPendiente) Then
                gCumplirServicio = 0
                Screen.MousePointer = 0
                MsgBox "La mercadería está pendiente de entrega. Verifique", vbInformation, "ATENCIÓN"
                EstadoEntregando False
            Else
                Foco tArticulo
            End If
        End If
        
    Else
        Screen.MousePointer = 0
        gDocumento = 0
        MsgBox "No existe un documento para las características ingresadas.", vbExclamation, "ATENCIÓN"
    End If
    
    RsAux.Close
    
    If idCliente <> 0 Then fnc_ControlComentariosAlEntregar (idCliente)
    If gCumplirServicio <> 0 Then AccionCumplirServicio gCumplirServicio
        
End Sub

Private Sub BuscoNota(Optional Tipo As Integer, Optional Sucursal As Long, Optional Serie As String, Optional Numero As Long, Optional Codigo As Long = 0)

    Screen.MousePointer = 11
    gCumplirServicio = 0
    
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
        
        lDocumento.Caption = UCase(NombreDocumento(RsAux!DocTipo) & " " & Trim(RsAux!DocSerie) & "-" & RsAux!DocNumero)
        
        CargoCliente RsAux!DocCliente
        
        'Cargo los articulos de la Tabla Devolucion-------------------------------------------------------------------------
        Dim aAEntregar As Integer: aAEntregar = 0
        lArticulo.ListItems.Clear

        Cons = "Select * from Devolucion, Articulo " & _
                  " Where DevCliente = " & RsAux!DocCliente & _
                  " And DevNota = " & RsAux!DocCodigo & _
                  " And DevLocal is Null" & _
                  " And DevArticulo = ArtID"
        Set rsXX = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsXX.EOF Then
            Do While Not rsXX.EOF
                'Si no es flete lo cargo
                If InStr(Fletes, rsXX!ArtID & ",") = 0 Then
                    Set itmX = lArticulo.ListItems.Add(Text:=Trim(rsXX!ArtCodigo))
                    itmX.Tag = rsXX!ArtID
                    
                    itmX.SubItems(1) = Trim(rsXX!ArtNombre)
                    itmX.SubItems(2) = rsXX!DevCantidad
                    aAEntregar = aAEntregar + rsXX!DevCantidad
                    
                    itmX.SubItems(3) = "0"
                    
                    If rsXX!DevCantidad = 0 Then itmX.SmallIcon = "no" Else: itmX.SmallIcon = "si"
                    If Not IsNull(rsXX!ArtBarCode) Then itmX.SubItems(4) = Trim(rsXX!ArtBarCode)
                    
                    'Si el Art requiere NroSerie el Key se arma con "S + id_Articulo", sino con "N + id_Articulo"
                    If rsXX!ArtNroSerie Then itmX.Key = "S" & rsXX!ArtID Else itmX.Key = "N" & rsXX!ArtID
                    
                End If
                rsXX.MoveNext
            Loop
            lAEntregar.Caption = "Total de Artículos a Recibir:      " & aAEntregar & " "
            EstadoEntregando True
        Else
            MsgBox "No existe registro de mercadería pendiente de ingreso al local para el documento seleccionado." & Chr(vbKeyReturn) & _
                        "Verifique si la mercadería fue recibida o que el documento sea el correcto.", vbExclamation, "No hay Mercadería Pendiente"
        End If
        rsXX.Close
            
        If RsAux!DocAnulado Then
            gCumplirServicio = 0
            Screen.MousePointer = 0
            MsgBox "El documento ingresado ha sido anulado. Verifique", vbCritical, "DOCUMENTO ANULADO"
            EstadoEntregando False
        Else
            If Not IsNull(RsAux!DocPendiente) Then
                gCumplirServicio = 0
                Screen.MousePointer = 0
                MsgBox "La mercadería está pendiente de entrega. Verifique", vbInformation, "ATENCIÓN"
                EstadoEntregando False
            Else
                Foco tArticulo
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
        gRemito = RsAux!RemCodigo
        gFechaDocumento = RsAux!DocFModificacion     'Siempre guardo la del Documento
        
        lDocumento.Caption = "REMITO Nº " & RsAux!DocNumero
        '-------------------------------------------------------------------
        
        CargoCliente RsAux!DocCliente
        
        Cons = "Select ArtID, ArtCodigo, ArtBarCode, ArtNombre, RReAEntregar ARetirar, ArtNroSerie, ArtTipo from RenglonRemito, Articulo" _
                & " Where RReRemito = " & RsAux!RemCodigo _
                & " And RReArticulo = ArtID "
            
        CargoArticulos Cons
        
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
                Foco tArticulo
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

'--------------------------------------------------------------------------------------------------------
'   Los campos seleccionados en la cosnulta se deben renombrar como:
'       ArtID, ArtCodigo, ArtBarCode, ArtNombre, ARetirar
'--------------------------------------------------------------------------------------------------------
Private Sub CargoArticulos(Consulta As String)

Dim aAEntregar As Integer

    On Error GoTo errCargar
    lArticulo.ListItems.Clear
    aAEntregar = 0
         
    Set rsXX = cBase.OpenResultset(Consulta, rdOpenDynamic, rdConcurValues)
    Do While Not rsXX.EOF
        'Si no es flete lo cargo
        If InStr(Fletes, rsXX!ArtID & ",") = 0 Then
            Set itmX = lArticulo.ListItems.Add(Text:=Trim(rsXX!ArtCodigo))
            itmX.Tag = rsXX!ArtID
            
            itmX.SubItems(1) = Trim(rsXX!ArtNombre)
            itmX.SubItems(2) = rsXX!ARetirar
            aAEntregar = aAEntregar + rsXX!ARetirar
            
            itmX.SubItems(3) = "0"
            
            If rsXX!ARetirar = 0 Then itmX.SmallIcon = "no" Else: itmX.SmallIcon = "si"
            
            If Not IsNull(rsXX!ArtBarCode) Then itmX.SubItems(4) = Trim(rsXX!ArtBarCode)
            
            itmX.SubItems(5) = Trim(rsXX!ArtTipo)
            
            'Si el Art requiere NroSerie el Key se arma con "S + id_Articulo", sino con "N + id_Articulo"
            If rsXX!ArtNroSerie Then itmX.Key = "S" & rsXX!ArtID Else itmX.Key = "N" & rsXX!ArtID
        End If
        
        rsXX.MoveNext
    Loop
    rsXX.Close
    
    lAEntregar.Caption = "Total de Artículos a Entregar:      " & aAEntregar & " "
    lAEntregar.Tag = aAEntregar
    
    If aAEntregar = 0 Then
        'Verifico si no es un Documento q Factura Servicios---------------------------------------------------------------
        Dim rsSer As rdoResultset
        Cons = "Select * From Servicio Where SerDocumento = " & gDocumento
        Set rsSer = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsSer.EOF Then
            Select Case rsSer!SerEstadoServicio
                Case EstadoS.Cumplido: MsgBox "El servicio " & rsSer!SerCodigo & ", al que está asociado la factura ya fue cumplido.", vbExclamation, "Factura de Servicios."
                Case EstadoS.Anulado: MsgBox "El servicio " & rsSer!SerCodigo & ", al que está asociado la factura fue anulado.", vbExclamation, "Factura de Servicios."
                Case EstadoS.Entrega:
                            If MsgBox("El servicio " & rsSer!SerCodigo & ", está pendiente de entrega a domicilio." & Chr(vbKeyReturn) & _
                                           "Si ud. lo cumple se cancelará la entrega. Desea entregarlo (cumplirlo)", vbQuestion + vbYesNo + vbDefaultButton2, "Factura de Servicios.") = vbYes Then gCumplirServicio = rsSer!SerCodigo
                Case Else: MsgBox "El servicio " & rsSer!SerCodigo & ", está asociado a esta factura." & Chr(vbKeyReturn) & _
                                           "Se va a cumplir el servicio y dar como entregado el producto", vbInformation, "Factura de Servicios."
                                           gCumplirServicio = rsSer!SerCodigo
            End Select
        
        Else
            MsgBox "El documento no tiene artículos para entregar." & Chr(vbKeyReturn) & "Están asociados a otro documento o ya se han entregado.", vbInformation, "ATENCIÓN"
        End If
        rsSer.Close
               
        EstadoEntregando False
    Else
        EstadoEntregando True
        Dim aRmv As Integer: aRmv = 1
        For I = 1 To lArticulo.ListItems.Count
            If aRmv > lArticulo.ListItems.Count Then Exit For
            If Val(lArticulo.ListItems(aRmv).SubItems(2)) = 0 Then
                lArticulo.ListItems.Remove aRmv
                lParcial.Visible = True
            Else
                aRmv = aRmv + 1
            End If
        Next
    End If
    
    Exit Sub

errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los artículos del documento.", Err.Description
End Sub

Private Sub EstadoEntregando(Estado As Boolean)

    tCBarra.Enabled = Not Estado
    tArticulo.Enabled = Estado
    
    If Estado Then
        tArticulo.BackColor = Blanco
        
        tCBarra.BackColor = Inactivo
    Else
        tArticulo.BackColor = Inactivo
        
        tCBarra.BackColor = Blanco
    End If
    
End Sub

Private Sub EntregoArticulo(Codigo As String, Cantidad As Currency, Optional EsBarCode As Boolean = True)

Dim sEntrego As Boolean
Dim aCodigo As String
    
    'Valido los parametros del lector y los codigos ingresados.-------------------------------------------------------------
    Dim bTieneBC As Boolean: bTieneBC = False
    For Each itmX In lArticulo.ListItems
        If EsBarCode Then aCodigo = Trim(itmX.SubItems(4)) Else aCodigo = Trim(itmX.Text)
        If Trim(aCodigo) = Trim(Codigo) Then
            If Trim(itmX.SubItems(4)) <> "" Then bTieneBC = True
            Exit For
        End If
    Next
    
    If Not InicioEMercaderia.bSinLector And bTieneBC Then
        If Not EsBarCode And Cantidad < 5 Then
            MsgBox "El artículo tiene código de barras, páselo por el lector para entregarlo.", vbExclamation, "Artículo con Código de Barras."
            Exit Sub
        End If
    End If
    
    If InicioEMercaderia.bSinLector And bTieneBC And EsBarCode Then
        MsgBox "El lector de barras está funcionando y ud. lo tiene deshabilitado." & Chr(vbKeyReturn) & "Acceda al menú lector para habilitarlo.", vbExclamation, "Habilite el Lector."
    End If
    
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
                        aNroSerie = InputBox("Ingrese el número de nerie del artículo entregado.", "(" & Trim(itmX.Text) & ") " & Trim(itmX.SubItems(1)))
                        If Trim(aNroSerie) <> "" Then
                            If Not arrAgregoElemento(itmX.Tag, Trim(aNroSerie)) Then aNroSerie = ""
                        End If
                    Loop
                End If
                '-------------------------------------------------------------------------------------------------------------------
            
            Else
                'No hay para entregar----------------------------------------------
                MsgBox "El / los " & Trim(itmX.SubItems(1)) & " ya se han entregado.", vbExclamation, "ATENCIÓN"
            End If
            
            sEntrego = True
            Exit For
        End If
    Next
    
    tArticulo.Text = ""
    
    If Not sEntrego Then
        MsgBox "El artículo ingresado no figura en la lista. No se debe entregar con éste documento.", vbCritical, "ATENCIÓN"
    End If
       
    'Si ya se entregaron TODOS Grabo------------------------------------------------
    sEntrego = True
    For Each itmX In lArticulo.ListItems
        If itmX.SmallIcon = "si" Then sEntrego = False: Exit For
    Next
    If sEntrego Then AccionGrabar
    '----------------------------------------------------------------------------------------
    
End Sub

Private Sub AccionGrabar()

Dim bDevolucion As Boolean      'Indica si es una devolucion por nota o documento comun

    'Valido si hay articulos entregados------------------------------------
    Dim sHay As Boolean
    sHay = False
    For Each itmX In lArticulo.ListItems
        If itmX.SubItems(3) <> 0 Then sHay = True: Exit For
    Next
    '--------------------------------------------------------------------------
    
    If Not sHay Then        'Accion Cancelar----------
        LimpioParaNuevaEntrega
        Exit Sub
    End If
        
    If gTipo = TipoDocumento.NotaCredito Or gTipo = TipoDocumento.NotaDevolucion Or gTipo = TipoDocumento.NotaEspecial Then
        bDevolucion = True
        'Como es una recepcion de devolucion de mercadería por nota, se deben devolver todos los articulos
        sHay = True
        For Each itmX In lArticulo.ListItems
            If Val(itmX.SubItems(2)) <> 0 Then sHay = False: Exit For
        Next
        If Not sHay Then
            MsgBox "Esto es una recepción de mercadería por devolución." & Chr(vbKeyReturn) & "El cliente debe devolver todos los artículos, de lo contrario no podrá realizar el ingreso.", vbExclamation, "Faltan Artículos"
            Foco tArticulo: Exit Sub
        End If
        
        If MsgBox("Confirma almacenar la recepción de la mercadería por devolución.", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
        
    Else
        bDevolucion = False
        If MsgBox("Confirma almacenar la entrega de mercadería.", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    End If
    
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
    
    If Not bDevolucion Then
        BorroProductosVendidos 'Por duplicaicones de códigos
        GraboDatosTablas
        GraboProductosVendidos gDocumento
    Else
        GraboDatosTablasDevolucion
        GraboProductosVendidos gDocumento, Alta:=False
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

    If gTipo = TipoDocumento.Remito Then aDocumento = gRemito Else aDocumento = gDocumento
    
    For Each itmX In lArticulo.ListItems
    
        If CCur(itmX.SubItems(3)) > 0 Then
            'Actualizo los datos en tabla Renlones------------------------------------------------------------------------------------
            If gTipo = TipoDocumento.Remito Then
                Cons = "Update RenglonRemito " _
                    & " Set RReAEntregar = RReAEntregar - " & CCur(itmX.SubItems(3)) _
                    & " Where RReRemito = " & aDocumento _
                    & " And RReArticulo = " & itmX.Tag
            Else
                Cons = "Update Renglon " _
                    & " Set RenARetirar = RenARetirar - " & CCur(itmX.SubItems(3)) _
                    & " Where RenDocumento = " & aDocumento _
                    & " And RenArticulo = " & itmX.Tag
            End If
            cBase.Execute Cons
            '-------------------------------------------------------------------------------------------------------------------------------
            
            If Val(itmX.SubItems(5)) <> paTipoArticuloServicio Then
                'Marco la Baja del STOCK AL LOCAL
                'Genero Movimiento
                MarcoMovimientoStockFisico CLng(Me.Tag), TipoLocal.Deposito, paCodigoDeSucursal, CLng(itmX.Tag), CCur(itmX.SubItems(3)), paEstadoArticuloEntrega, -1, gTipo, aDocumento
                'Bajo del Stock en Local
                MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, CLng(itmX.Tag), CCur(itmX.SubItems(3)), paEstadoArticuloEntrega, -1
                
                'Marco el Movimiento del STOCK VIRTUAL
                'Genero Movimiento
                MarcoMovimientoStockEstado CLng(Me.Tag), CLng(itmX.Tag), CCur(itmX.SubItems(3)), TipoMovimientoEstado.ARetirar, -1, gTipo, aDocumento, paCodigoDeSucursal
                'Bajo del Stock Total
                MarcoMovimientoStockTotal CLng(itmX.Tag), TipoEstadoMercaderia.Virtual, TipoMovimientoEstado.ARetirar, CCur(itmX.SubItems(3)), -1
            End If
        End If
    Next

End Sub

Private Sub GraboDatosTablasDevolucion()

Dim aDocumento As Long
Dim rsDev As rdoResultset
    
    For Each itmX In lArticulo.ListItems
    
        If CCur(itmX.SubItems(3)) > 0 Then
            'Actualizo los datos en tabla Devoluciones------------------------------------------------------------------------------------
            Cons = "Select * From Devolucion" & _
                    " Where DevNota = " & gDocumento & _
                    " And DevArticulo = " & itmX.Tag & _
                    " And DevLocal is Null"
            Set rsDev = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            rsDev.Edit
            rsDev!DevLocal = paCodigoDeSucursal
            rsDev!DevFAltaLocal = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
            rsDev.Update: rsDev.Close
                
            '-------------------------------------------------------------------------------------------------------------------------------
            
            'Marco el ALTA del STOCK AL LOCAL
            'Genero Movimiento
            MarcoMovimientoStockFisico CLng(Me.Tag), TipoLocal.Deposito, paCodigoDeSucursal, CLng(itmX.Tag), CCur(itmX.SubItems(3)), paEstadoArticuloEntrega, 1, gTipo, gDocumento
            'Alta del Stock en Local
            MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, CLng(itmX.Tag), CCur(itmX.SubItems(3)), paEstadoArticuloEntrega, 1
            
            'Marco el Movimiento del STOCK VIRTUAL (No van porque ya los hizo la nota)
            'Genero Movimiento
            'MarcoMovimientoStockEstado CLng(Me.Tag), CLng(itmX.Tag), CCur(itmX.SubItems(3)), TipoMovimientoEstado.ARetirar, -1, gTipo, aDocumento, paCodigoDeSucursal
            
            'Sumo al Stock Total
            MarcoMovimientoStockTotal CLng(itmX.Tag), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, CCur(itmX.SubItems(3)), 1
        End If
    Next

End Sub

Private Sub GraboProductosVendidos(idDocumento As Long, Optional Alta As Boolean = True)

    Dim rsPV As rdoResultset
    If Alta Then
        Cons = "Select * from ProductosVendidos Where PVeDocumento = " & idDocumento
        Set rsPV = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        For I = LBound(arrNroSerie) To UBound(arrNroSerie)
            If arrNroSerie(I).Articulo <> -1 And Trim(arrNroSerie(I).NroSerie) <> "" Then
                rsPV.AddNew
                rsPV!PVeDocumento = idDocumento
                rsPV!PVeArticulo = arrNroSerie(I).Articulo
                rsPV!PVeNSerie = Trim(arrNroSerie(I).NroSerie)
                rsPV.Update
            End If
        Next
        
        rsPV.Close
    
    Else
        'Como es una devolucion gDocumento = al IDNOTA --> busco el doc de la nota
        Dim idFactura As Long
        Cons = "Select * from Nota Where NotNota = " & idDocumento
        Set rsPV = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsPV.EOF Then idFactura = rsPV!NotFactura
        rsPV.Close
        
        For I = LBound(arrNroSerie) To UBound(arrNroSerie)
            If arrNroSerie(I).Articulo <> -1 And Trim(arrNroSerie(I).NroSerie) <> "" Then
                
                Cons = "Select * from ProductosVendidos Where PVeDocumento = " & idFactura & _
                            " And PVeArticulo = " & arrNroSerie(I).Articulo & _
                            " And PVeNSerie = '" & Trim(arrNroSerie(I).NroSerie) & "'"
                Set rsPV = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not rsPV.EOF Then rsPV.Delete
                rsPV.Close
        
            End If
        Next
    End If
    
End Sub

Private Function BorroProductosVendidos()

    Dim rsPV As rdoResultset
            
    For I = LBound(arrNroSerie) To UBound(arrNroSerie)
        If arrNroSerie(I).Articulo <> -1 And Trim(arrNroSerie(I).NroSerie) <> "" Then
            
            Cons = "Select * from ProductosVendidos " & _
                    " Where PVeArticulo = " & arrNroSerie(I).Articulo & _
                    " And PVeNSerie = '" & Trim(arrNroSerie(I).NroSerie) & "'"
            Set rsPV = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rsPV.EOF Then rsPV.Delete
            rsPV.Close
        End If
    Next
    
End Function


Private Sub LimpioParaNuevaEntrega(Optional Minimizar As Boolean = True)

    EstadoEntregando False
    lAEntregar.Caption = ""
    lArticulo.ListItems.Clear
    tCBarra.Text = ""
    lDocumento.Caption = ""
        
    lCiRuc.Caption = "N/D"
    lTitular.Caption = "N/D"
    
    Foco tCBarra
    gDocumento = 0
    gRemito = 0
    lParcial.Visible = False
    
    ReDim arrNroSerie(0)
    
    If Minimizar Then Me.WindowState = vbMinimized
    
End Sub

'----------------------------------------------------------------------------------
'   Interpreta el Texto del Codigo de Barras
'   Formato:    XDXXXX          TipoDocumento   D Numero de Documento
'----------------------------------------------------------------------------------
Private Sub FormatoBarras(Texto As String)

Dim aCodDoc As Long
    
    On Error GoTo errInt
    Texto = UCase(Texto)
    
    '1) Veo si es x codigo de barras o x ids de documento
    If (Mid(Texto, 2, 1) = "D" And IsNumeric(Mid(Texto, 1, 1)) And Len(Texto) > 3) Or (Mid(Texto, 3, 1) = "D" And IsNumeric(Mid(Texto, 1, 2)) And Len(Texto) > 4) Then   'Codigo de Barras
        gTipo = CLng(Mid(Texto, 1, InStr(Texto, "D") - 1))
        aCodDoc = CLng(Trim(Mid(Texto, InStr(Texto, "D") + 1, Len(Texto))))
    Else
        'Puso Serie y Numero de Documento o Numero de Remito
        If IsNumeric(Texto) Then
            gTipo = TipoDocumento.Remito        'Remito
            aCodDoc = Texto
        Else
            'Puso Serie y Numero de Documento
             If Not zfn_BuscoDocPorTexto(Texto, aCodDoc, gTipo) Then aCodDoc = -1
        End If
        
    End If
    
    If aCodDoc = -1 Then
        MsgBox "No existe un documento que coincida con los valores ingresados.", vbExclamation, "No hay Datos"
        Exit Sub
    End If
    
    Select Case gTipo
        Case TipoDocumento.Remito:  BuscoRemito aCodDoc
        
        Case TipoDocumento.Contado, TipoDocumento.Credito: BuscoDocumento Tipo:=gTipo, Codigo:=aCodDoc
        
        Case TipoDocumento.NotaCredito, TipoDocumento.NotaDevolucion, TipoDocumento.NotaEspecial: BuscoNota Tipo:=gTipo, Codigo:=aCodDoc
        
        Case Else:  MsgBox "El código de barras ingresado no es correcto. El documento no coincide con los predefinidos.", vbCritical, "ATENCIÓN"
    End Select
    Screen.MousePointer = 0
    Exit Sub
    
errInt:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al interpretar el código de barras.", Err.Description
End Sub


Private Function zfn_BuscoDocPorTexto(adTexto As String, retIDDoc As Long, retIDTipoD) As Boolean
On Error GoTo errDoc
    zfn_BuscoDocPorTexto = False
    
    Dim mDSerie As String, mDNumero As Long
    Dim adQ As Integer, adCodigo As Long, adTipoD As Integer
        
    If InStr(adTexto, "-") <> 0 Then
        mDSerie = Mid(adTexto, 1, InStr(adTexto, "-") - 1)
        mDNumero = Val(Mid(adTexto, InStr(adTexto, "-") + 1))
    Else
        mDSerie = Mid(adTexto, 1, 1)
        mDNumero = Val(Mid(adTexto, 2))
    End If
    
    adTexto = UCase(mDSerie) & "-" & mDNumero
        
    Screen.MousePointer = 11
    adQ = 0: adTexto = ""
    
    'Cargo combo con tipos de docuemento--------------------------------------
    Cons = "Select DocCodigo, DocTipo, DocFecha as Fecha, DocSerie as Serie, Convert(char(7),DocNumero) as Numero " & _
               " From Documento " & _
               " Where DocSerie = '" & mDSerie & "'" & _
               " And DocNumero = " & mDNumero & _
               " And DocTipo IN (" & TipoDocumento.Contado & ", " & TipoDocumento.Credito & ", " & TipoDocumento.NotaCredito & ", " & _
                                               TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial & ", " & TipoDocumento.Remito & ")" & _
                " And DocAnulado = 0"
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        adCodigo = RsAux!DocCodigo
        adTipoD = RsAux!DocTipo
        adQ = 1
        RsAux.MoveNext: If Not RsAux.EOF Then adQ = 2
    End If
    RsAux.Close
        
        Select Case adQ
            Case 2
                Dim miLDocs As New clsListadeAyuda
                If miLDocs.ActivarAyuda(cBase, Cons, 4100, 2) <> 0 Then
                    adCodigo = miLDocs.RetornoDatoSeleccionado(0)
                    adTipoD = miLDocs.RetornoDatoSeleccionado(1)
                End If
                Set miLDocs = Nothing
                Me.Refresh
        End Select
        
        If adCodigo > 0 Then
            'lDoc.Tag = adCodigo: lDoc.Caption = adTexto
            zfn_BuscoDocPorTexto = True
            retIDDoc = adCodigo
            retIDTipoD = adTipoD
        Else
            'lDoc.Caption = " No Existe !!"
            zfn_BuscoDocPorTexto = False
        End If
        
        Screen.MousePointer = 0
    'End If
    
    Exit Function
errDoc:
    clsGeneral.OcurrioError "Error al buscar el documento.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub AccionCumplirServicio(Codigo As Long)

    On Error GoTo errCumplir
    Dim rsSer As rdoResultset
    Screen.MousePointer = 11
    FechaDelServidor
    
    'Tabla Servicio-----------------------------------------------------------------------------------------------------
    Cons = "Select * from Servicio Where SerCodigo = " & Codigo
    Set rsSer = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    rsSer.Edit
    
    If IsNull(rsSer!SerComentarioR) Then rsSer!SerComentarioR = "Cumplido y Entregado por sistema E.M."
    rsSer!SerFCumplido = Format(gFechaServidor, "mm/dd/yyyy")
    rsSer!SerModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    rsSer!SerUsuario = Val(Me.Tag)
    rsSer!SerEstadoServicio = EstadoS.Cumplido
    rsSer.Update: rsSer.Close
    
    LimpioParaNuevaEntrega
    Screen.MousePointer = 0
    Exit Sub
    
errCumplir:
    clsGeneral.OcurrioError "Ocurrió un error al cumplir el servicio Nº " & Codigo, Err.Description
    Screen.MousePointer = 0
End Sub

Private Function arrAgregoElemento(aIdArticulo As Long, aSerie As String) As Boolean
    
    On Error GoTo errAgregar
    arrAgregoElemento = False
    If arrBuscoElemento(aIdArticulo, aSerie) <> 0 Then
        MsgBox "El nro. de serie ingresado ya fue entregado !!!!." & vbCrLf & vbCrLf & "Nº Serie: " & Trim(aSerie), vbExclamation, "Artículo Entregado"
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
        MsgBox "Este artículo no fue entregado !!!!." & vbCrLf & vbCrLf & "Nº Serie: " & Trim(aSerie), vbExclamation, "Artículo NO Entregado"
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

