VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIngDe 
   BackColor       =   &H00DEEDEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Mercadería por Devoluciones"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIngDe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frDetalle 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEEDEF&
      Caption         =   "Detalle de Ingreso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      TabIndex        =   28
      Top             =   2520
      Width           =   5415
      Begin VB.ComboBox tArticulo 
         Height          =   315
         Left            =   900
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox tCantidad 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4620
         MaxLength       =   4
         TabIndex        =   9
         Top             =   240
         Width           =   555
      End
      Begin VB.CheckBox chEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEEDEF&
         Caption         =   "Con la Manguera"
         Enabled         =   0   'False
         ForeColor       =   &H00CC5000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.OptionButton obIngreso 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEEDEF&
      Caption         =   "Ingreso relacionado al cliente  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   205
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   3255
   End
   Begin VB.OptionButton obIngreso 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEEDEF&
      Caption         =   "Ingresa relacionado al Documento  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   205
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   3615
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   5700
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7355
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Picture         =   "frmIngDe.frx":08CA
            Key             =   "printer"
         EndProperty
      EndProperty
   End
   Begin vsViewLib.vsPrinter vsFicha 
      Height          =   1935
      Left            =   5280
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   3795
      _Version        =   196608
      _ExtentX        =   6694
      _ExtentY        =   3413
      _StockProps     =   229
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ConvInfo        =   1418783674
      PageBorder      =   0
      BackColor       =   -2147483633
   End
   Begin VB.CommandButton bSalir 
      Caption         =   "&Salir"
      Height          =   315
      Left            =   4560
      TabIndex        =   23
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton bCancelar 
      Caption         =   "Canc&elar"
      Height          =   315
      Left            =   3480
      TabIndex        =   21
      Top             =   5280
      Width           =   975
   End
   Begin MSMask.MaskEdBox tRuc 
      Height          =   285
      Left            =   2760
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99 999 999 9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox tCi 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#.###.###-#"
      PromptChar      =   "_"
   End
   Begin VB.TextBox tCBarra 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox tComentario 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      MaxLength       =   60
      TabIndex        =   15
      Top             =   4920
      Width           =   4455
   End
   Begin VB.TextBox tUsuario 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   19
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      Height          =   195
      Left            =   180
      TabIndex        =   27
      Top             =   2115
      Width           =   615
   End
   Begin VB.Label lComentDoc 
      BackStyle       =   0  'Transparent
      Caption         =   "Ctdo. B 5460"
      ForeColor       =   &H00CC6900&
      Height          =   255
      Left            =   1200
      TabIndex        =   26
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Comentario:"
      Height          =   255
      Left            =   180
      TabIndex        =   25
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lFechaDoc 
      BackStyle       =   0  'Transparent
      Caption         =   "Ctdo. B 5460"
      ForeColor       =   &H00CC6900&
      Height          =   255
      Left            =   1200
      TabIndex        =   24
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Emisión:"
      Height          =   255
      Left            =   180
      TabIndex        =   22
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      Height          =   255
      Left            =   180
      TabIndex        =   20
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lClienteDoc 
      BackStyle       =   0  'Transparent
      Caption         =   "Juan Alberto"
      ForeColor       =   &H00CC6900&
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label lDocumento 
      BackStyle       =   0  'Transparent
      Caption         =   "Ctdo. B 5460"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2640
      MouseIcon       =   "frmIngDe.frx":09DC
      MousePointer    =   99  'Custom
      TabIndex        =   16
      ToolTipText     =   "Click Accede a Detalle de Factura"
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&CI/RUC:"
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   1815
      Width           =   855
   End
   Begin VB.Label lTitular 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "111111111111111"
      ForeColor       =   &H00CC6900&
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   2160
      UseMnemonic     =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Documento:"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lComentario 
      BackStyle       =   0  'Transparent
      Caption         =   "Co&mentario:"
      Height          =   255
      Left            =   180
      TabIndex        =   13
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario:"
      Height          =   255
      Left            =   180
      TabIndex        =   17
      Top             =   5280
      Width           =   615
   End
End
Attribute VB_Name = "frmIngDe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TipoSuceso
    DiferenciaDeArticulos = 11
End Enum

Private Enum TipoLocal
    Camion = 1
    Deposito = 2
End Enum

Private Enum TipoDocumento
    Contado = 1
    Credito = 2
    NotaDevolucion = 3
    NotaCredito = 4
    Remito = 6
    NotaEspecial = 10
    Devolucion = 28
End Enum

Private Enum TipoCliente
    Cliente = 1
    Empresa = 2
End Enum

'------------------------------------------------------------------
Private Type typArtEnCombo
    Articulo As Long
    QQuePuede As Integer
    NroSerie As Boolean
End Type

Dim arrArtCombo() As typArtEnCombo
Dim arrNroSerie() As String
'------------------------------------------------------------------
Private Fletes As String

Private Sub bCancelar_Click()
    ctrl_CleanCliente
    ctrl_CleanDocumento
    ReDim arrNroSerie(0)
    ReDim arrArtCombo(0)
    tArticulo.Clear
    tCantidad.Text = ""
    tComentario.Text = ""
    With tUsuario: .Text = "": .Tag = "": End With
    Foco tCBarra
End Sub

Private Sub bSalir_Click()
    Unload Me
End Sub

Private Sub chEstado_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tComentario
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0: Me.Refresh
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    
    ObtengoSeteoForm Me
    Me.Width = 5745
    
    lDocumento.Font.Underline = True
    Status.Panels("printer").Text = paPrintConfD
    
    ctrl_CleanCliente
    ctrl_CleanDocumento
    s_SetCtrlOpt
    
    With vsFicha
        .Orientation = orPortrait:
        .PaperSize = paPrintConfPaperSize
    End With
    FechaDelServidor
    Fletes = CargoArticulosDeFlete
    
    On Error Resume Next
    Cons = "Select Top 1 * From Devolucion"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If RsAux.rdoColumns("DevComentario").Size > 0 Then
            tComentario.MaxLength = RsAux.rdoColumns("DevComentario").Size
        End If
    End If
    RsAux.Close
    
    'Estados artículo devolución ......................................................................
    Cons = "Select * From EstadoArticuloDevolucion Where EADHabilitado <> 0 Order By EADOrden"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        chEstado(0).Visible = False
        frDetalle.Height = tArticulo.Top + tArticulo.Height + (tArticulo.Top / 2)
    Else
        With chEstado(0)
            .Tag = RsAux!EADCodigo
            .Caption = Trim(RsAux!EADTexto)
        End With
        RsAux.MoveNext
        Do While Not RsAux.EOF
            Load chEstado(chEstado.UBound + 1)
            With chEstado(chEstado.UBound)
                .Visible = True
                .Tag = RsAux!EADCodigo
                .Caption = Trim(RsAux!EADTexto)
                
                If chEstado.UBound Mod 2 <> 0 Then
                    .Left = 2760
                Else
                    .Left = 120
                End If
                .Top = chEstado(chEstado.UBound - 2).Top + chEstado(chEstado.UBound - 2).Height + 60
            End With
            RsAux.MoveNext
        Loop
        frDetalle.Height = chEstado(chEstado.UBound).Top + chEstado(chEstado.UBound).Height + 120
    End If
    RsAux.Close
    
    lComentario.Top = frDetalle.Top + frDetalle.Height + 60
    tComentario.Top = lComentario.Top
    With lUsuario
        .Top = lComentario.Top + tComentario.Height + 60
        tUsuario.Top = .Top
        bCancelar.Top = .Top
        bSalir.Top = .Top
    End With
    
    Me.Height = bSalir.Top + bSalir.Height + 480 + Status.Height
    
    'Estados artículo devolución ......................................................................
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al cargar el formulario.", Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    GuardoSeteoForm Me
    CierroConexion
    Set clsGeneral = Nothing
    Erase arrArtCombo
    Erase arrNroSerie
    End
End Sub

Private Sub Label1_Click()
    Foco tCBarra
End Sub

Private Sub Label2_Click()
    Foco tArticulo
End Sub

Private Sub lComentario_Click()
    Foco tComentario
End Sub

Private Sub lDocumento_Click()
    If Val(lDocumento.Tag) > 0 Then EjecutarApp App.Path & "\Detalle de Factura.exe", Val(lDocumento.Tag)
End Sub

Private Sub lUsuario_Click()
    Foco tUsuario
End Sub

Private Sub obIngreso_Click(Index As Integer)
    If Index = 0 Then
        ctrl_CleanCliente
    Else
        ctrl_CleanDocumento
    End If
    s_SetCtrlOpt
End Sub

Private Sub obIngreso_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If obIngreso(0).Value Then tCBarra.SetFocus Else tCi.SetFocus
    End If
End Sub

Private Sub Status_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
    If "printer" = Panel.Key Then
        prj_GetPrinter True
        Panel.Text = paPrintConfD
    End If
End Sub

Private Sub tArticulo_Change()
    If obIngreso(1).Value Then
        If tArticulo.ListCount > 0 Then tArticulo.RemoveItem 0
    End If
End Sub

Private Sub tArticulo_GotFocus()
    With tArticulo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrTA
Dim fModificacion As Date
Dim sNombre As String, lID As Long
    
    If KeyCode = vbKeyReturn Then
    
        If Trim(tArticulo.Text) <> "" Then
            
            If obIngreso(1).Value And Val(lTitular.Tag) = 0 Then Exit Sub
            
            'No busco
            If obIngreso(0).Value Then
                If tArticulo.ListIndex <> -1 Then Foco tCantidad
                Exit Sub
            Else
                If Val(tArticulo.Tag) > 0 Then Foco tCantidad: Exit Sub
            End If
                
            If Not IsNumeric(tArticulo.Text) Then   'Busqueda por nombre
                
                Cons = "Select ArtID, 'Código' = ArtCodigo , 'Nombre' = ArtNombre From Articulo " _
                        & " Where ArtNombre Like '" & Replace(tArticulo.Text, " ", "%") & "%' Order by ArtNombre"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
                If RsAux.EOF Then
                    RsAux.Close
                    MsgBox "No se encontró un artículo para los datos ingresados.", vbInformation, "ATENCIÓN"
                Else
                    RsAux.MoveNext
                    If RsAux.EOF Then
                        RsAux.MoveFirst
                        sNombre = Trim(RsAux!Nombre)
                        lID = RsAux!ArtID
                        RsAux.Close
                    Else
                        RsAux.Close
                        Dim objAyuda As New clsListadeAyuda
                        If objAyuda.ActivarAyuda(cBase, Cons, 5000, 1, "Artículos") > 0 Then
                            sNombre = Trim(objAyuda.RetornoDatoSeleccionado(2))
                            lID = objAyuda.RetornoDatoSeleccionado(0)
                        End If
                        Set objAyuda = Nothing
                    End If
                End If
            Else                                            'Busqueda por codigo
                Cons = "Select ArtID, 'Código' = ArtCodigo, 'Nombre' = ArtNombre From Articulo Where ArtCodigo = " & Val(tArticulo.Text)
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
                If RsAux.EOF Then
                    RsAux.Close
                    MsgBox "No se encontró un artículo para el código ingresado.", vbInformation, "ATENCIÓN"
                Else
                    sNombre = Trim(RsAux!Nombre)
                    lID = RsAux!ArtID
                End If
            End If
            If lID > 0 Then
                If InStr(1, "," & paArtsNoNotaEsp & ",", "," & lID & ",") > 0 Then
                    MsgBox "Este artículo está expresamente inhabilitado para emitir fichas de devolución, por favor consulte.", vbExclamation, "Atención"
                    sNombre = ""
                    lID = 0
                Else
                    With tArticulo
                        .Clear
                        .AddItem sNombre
                        .ItemData(.NewIndex) = lID
                        .ListIndex = 0
                    End With
                    Foco tCantidad
                End If
            End If
        Else
            Foco tComentario
        End If
    End If
    Exit Sub
    
ErrTA:
    clsGeneral.OcurrioError "Ocurrió un error al cargar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tArticulo_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub tCantidad_Change()
    ReDim arrNroSerie(0)
End Sub

Private Sub tCantidad_GotFocus()
    With tCantidad
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Status.SimpleText = " Ingrese la cantidad de artículos que devuelve."
End Sub

Private Sub tCantidad_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If tArticulo.ListIndex = -1 Then Foco tArticulo: Exit Sub
        If Not IsNumeric(tCantidad.Text) Then MsgBox "Debe ingresar un nro. positivo.", vbExclamation, "ATENCIÒN": Exit Sub
        If obIngreso(1).Value Then
            If Val(lTitular.Tag) = 0 Then MsgBox "No hay seleccionado un cliente.", vbExclamation, "ATENCIÓN": Exit Sub
            If CantidadArticulosEnDevolucion(Val(lTitular.Tag), 0, tArticulo.ItemData(0)) > 0 Then
                    MsgBox "Ya existe un ingreso de devolución pendiente para el artículo y el cliente seleccionado, verifique.", vbInformation, "ATENCIÓN"
            End If
        End If
        If tArticulo.ListIndex <> -1 And Val(tCantidad.Text) > 0 Then
            If chEstado(0).Visible Then
                chEstado(0).SetFocus
            Else
                Foco tComentario
            End If
        End If
    End If
    
End Sub

Private Sub tCBarra_Change()
    
    If Val(lTitular.Tag) > 0 Then ctrl_CleanCliente
    If Val(lDocumento.Tag) > 0 Then ctrl_CleanDocumento
    ReDim arrNroSerie(0)
    tArticulo.Clear
    ReDim arrArtCombo(0)
    ctrl_SetArticulo False
    
End Sub

Private Sub tCBarra_GotFocus()
    With tCBarra
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tCBarra_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tCBarra.Text = "") Then tCi.SetFocus: Exit Sub
        If Val(lDocumento.Tag) > 0 And tArticulo.ListCount > 0 Then tArticulo.SetFocus: Exit Sub
        Screen.MousePointer = 11
        s_SetDocumento
        Foco tCBarra
        Screen.MousePointer = 0
    End If
    
End Sub

Private Sub tCi_Change()
    
    If Val(lDocumento.Tag) > 0 Then ctrl_CleanDocumento
    If Val(lTitular.Tag) > 0 Then ctrl_SetArticulo False: lTitular.Caption = "": lTitular.Tag = "": tRuc.Tag = "": tRuc.Text = ""
    ReDim arrNroSerie(0)
    tArticulo.Clear
    ReDim arrArtCombo(0)
    
End Sub

Private Sub tCi_GotFocus()
    tCi.SelStart = 0: tCi.SelLength = 11
    Status.SimpleText = " Ingrese la cédula del cliente.([F2]=Ficha, [F3]=Nuevo, [F4]=Buscar) "
End Sub

Private Sub tCi_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift <> 0 Then Exit Sub
    Select Case KeyCode
        Case vbKeyF2: If Val(tRuc.Tag) = "" Then FichaCliente TipoCliente.Cliente
        Case vbKeyF3: NuevoCliente TipoCliente.Cliente
        Case vbKeyF4: BuscarClientes TipoCliente.Cliente
    End Select
End Sub

Private Sub TCI_KeyPress(KeyAscii As Integer)
Dim lCli As Long
    
    If KeyAscii = vbKeyReturn Then
        If Val(lTitular.Tag) > 0 Then
           tArticulo.SetFocus
           Exit Sub
        End If
        
        Dim aCi As String
        'Valido la Cédula ingresada----------
        If Trim(tCi.Text) <> "" Then
            If Len(tCi.Text) <> 8 Then
                MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
            If Not clsGeneral.CedulaValida(tCi.Text) Then
                MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
        End If
        Screen.MousePointer = 11
        'Busco el Cliente -----------------------
        If Trim(tCi.Text) <> "" Then
            lCli = BuscoClienteCIRUC(tCi.Text)
            If lCli = 0 Then
                Screen.MousePointer = 0
                MsgBox "No existe un cliente para la cédula ingresada.", vbExclamation, "ATENCIÓN"
            Else
                 db_FindCliente lCli, lTitular, False
            End If
            Call tCi_GotFocus
        Else
            tRuc.SetFocus
        End If
        Screen.MousePointer = 0
    End If

End Sub

Private Sub tCi_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub tComentario_GotFocus()
    With tComentario
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tUsuario
End Sub

Private Sub tRuc_Change()
    If Val(lDocumento.Tag) > 0 Then ctrl_CleanDocumento
    If Val(lTitular.Tag) > 0 Then ctrl_SetArticulo False: lTitular.Tag = "": lTitular.Caption = "": tCi.Tag = "": tCi.Text = ""
End Sub

Private Sub tRuc_GotFocus()
    tRuc.SelStart = 0: tRuc.SelLength = 15
    Status.SimpleText = " Ingrese el R.U.C. de la empresa.([F2]=Ficha,[F3]=Nuevo,[F4]=Buscar) "
End Sub

Private Sub tRuc_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift <> 0 Then Exit Sub
    Select Case KeyCode
        Case vbKeyF2:  If Val(tRuc.Tag) = 2 Then FichaCliente TipoCliente.Empresa
        Case vbKeyF3: NuevoCliente TipoCliente.Empresa
        Case vbKeyF4: BuscarClientes TipoCliente.Empresa
    End Select
End Sub

Private Sub tRuc_KeyPress(KeyAscii As Integer)
Dim lCli As Long
    
    If KeyAscii = vbKeyReturn Then
        If Val(lTitular.Tag) > 0 Then
           tArticulo.SetFocus
           Exit Sub
        End If
        
        If Trim(tRuc.Text) <> "" Then
            Screen.MousePointer = 11
            lCli = BuscoClienteCIRUC(Trim(tRuc.Text))
            If lCli = 0 Then
                Screen.MousePointer = 0
                MsgBox "No existe un cliente para el número de RUC ingresado.", vbExclamation, "ATENCIÓN"
            Else
                'Cargo Datos del Cliente Seleccionado------------------------------------------------
                 db_FindCliente lCli, lTitular, False
            End If
            Call tRuc_GotFocus
        Else
            tCi.SetFocus
        End If
        Screen.MousePointer = 0
    End If
    
End Sub

Private Sub tRuc_LostFocus()
    Status.SimpleText = ""
End Sub

Private Function BuscoClienteCIRUC(CiRuc As String)
    On Error GoTo errBuscar
    BuscoClienteCIRUC = 0
    Cons = "Select * from Cliente Where CliCiRuc = '" & Trim(CiRuc) & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then BuscoClienteCIRUC = RsAux!CliCodigo
    RsAux.Close
    Exit Function
errBuscar:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el cliente."
    Screen.MousePointer = 0
End Function

Public Sub BuscoClienteSeleccionado(Codigo As Long)
Dim aCliente As Long
    Screen.MousePointer = 11
    ctrl_CleanCliente
    If Codigo > 0 Then db_FindCliente Codigo, lTitular, False
    Screen.MousePointer = 0
    Exit Sub
errSolicitud:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos de la solicitud.", Err.Description
End Sub

Private Sub FichaCliente(Optional iTipo As Integer = 0)
On Error GoTo ErrFC
Dim lID As Long
    
    If Val(lTitular.Tag) = 0 Then Exit Sub
    Screen.MousePointer = 11
    Dim objCliente As New clsCliente
    If iTipo = 0 Then
        If Val(tCi.Tag) = 1 Then iTipo = 1 Else iTipo = 2
    End If
    
    If iTipo > 0 Then
        objCliente.Personas Val(lTitular.Tag), 0, 0
    Else
        objCliente.Empresas Val(lTitular.Tag), False
    End If
    Me.Refresh
    
    lID = objCliente.IDIngresado
    Set objCliente = Nothing
    If lID <> 0 Then
        db_FindCliente lID, lTitular, False
        If iTipo = 1 Then Call tCi_GotFocus Else Call tRuc_GotFocus
    Else
        ctrl_CleanCliente
    End If
    Screen.MousePointer = 0
    Exit Sub
    
ErrFC:
    clsGeneral.OcurrioError "Ocurrió un error al ir a ficha de cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub NuevoCliente(aTipoCliente As Integer)
On Error GoTo ErrFC
    Dim lIDCli As Long
    
    Screen.MousePointer = 11
    ctrl_CleanCliente
    Dim objCliente As New clsCliente
    If aTipoCliente = TipoCliente.Cliente Then
        objCliente.Personas 0, 0, 1
    Else
        objCliente.Empresas 0, True
    End If
    Me.Refresh
    lIDCli = objCliente.IDIngresado
    Set objCliente = Nothing
    If lIDCli > 0 Then
        db_FindCliente lIDCli, lTitular, False
        If Val(tCi.Tag) = 1 Then tCi.SetFocus Else Call tRuc_GotFocus
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrFC:
    clsGeneral.OcurrioError "Ocurrió un error al ir a ficha de cliente.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub BuscarClientes(aTipoCliente As Integer)
    
    Screen.MousePointer = 11
    Dim objBuscar As New clsBuscarCliente
    Dim aTipo As Integer, aCliente As Long
    
    ctrl_CleanCliente
    If aTipoCliente = TipoCliente.Cliente Then objBuscar.ActivoFormularioBuscarClientes cBase, Persona:=True
    If aTipoCliente = TipoCliente.Empresa Then objBuscar.ActivoFormularioBuscarClientes cBase, Empresa:=True
    Me.Refresh
    aTipo = objBuscar.BCTipoClienteSeleccionado
    aCliente = objBuscar.BCClienteSeleccionado
    Set objBuscar = Nothing
    
    On Error GoTo errCargar
    If aCliente <> 0 Then
        db_FindCliente aCliente, lTitular, False
        If aTipo = 1 Then
            Call tCi_GotFocus
        Else
            Call tRuc_GotFocus
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function f_FormatoBarras() As Long
Dim aCodDoc As Long
Dim gTipo As Integer

    On Error GoTo errInt
    f_FormatoBarras = 0
    
    gTipo = CLng(Mid(tCBarra.Text, 1, InStr(tCBarra.Text, "D") - 1))
    aCodDoc = CLng(Trim(Mid(tCBarra.Text, InStr(tCBarra.Text, "D") + 1, Len(tCBarra.Text))))
    
    Select Case gTipo
        Case TipoDocumento.Contado, TipoDocumento.Credito, _
            TipoDocumento.NotaCredito, TipoDocumento.NotaDevolucion, TipoDocumento.NotaEspecial
            f_FormatoBarras = aCodDoc
        Case Else:  MsgBox "El código de barras ingresado no es correcto. El documento no coincide con los predefinidos (contado ó crédito).", vbCritical, "ATENCIÓN"
    End Select
    Exit Function
errInt:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al interpretar el código de barras.", Err.Description
End Function

Private Sub db_FindArticuloNota()
Dim RsXX As rdoResultset, aValor As Long

    Screen.MousePointer = 11
    
    'Cargo los articulos de la Tabla Devolucion-------------------------------------------------------------------------
    ReDim arrNroSerie(0)
    tArticulo.Clear
    ReDim arrArtCombo(0)

    Cons = "Select * from Devolucion, Articulo " & _
              " Where DevCliente = " & Val(lClienteDoc.Tag) & _
              " And DevNota = " & Val(lDocumento.Tag) & _
              " And DevLocal is Null And DevArticulo = ArtID"
              
    Set RsXX = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsXX.EOF Then
    
        Do While Not RsXX.EOF
            'Si no es Flete lo Cargo
            If InStr(Fletes, RsXX!ArtID & ",") = 0 Then
                With tArticulo
                    .AddItem Format(RsXX!ArtCodigo, "(#,000,000)") & " " & Trim(RsXX!ArtNombre)
                    .ItemData(.NewIndex) = RsXX!ArtID
                End With
                If arrArtCombo(0).Articulo > 0 Then ReDim Preserve arrArtCombo(UBound(arrArtCombo) + 1)
                With arrArtCombo(UBound(arrArtCombo))
                    .Articulo = RsXX!ArtID
                    .QQuePuede = RsXX!DevCantidad
                    .NroSerie = RsXX!ArtNroSerie
                End With
            End If
            RsXX.MoveNext
        Loop
    Else
        MsgBox "No existe registro de mercadería pendiente de ingreso al local para el documento seleccionado." & Chr(vbKeyReturn) & _
                    "Verifique si la mercadería fue recibida o que el documento sea el correcto.", vbExclamation, "No hay Mercadería Pendiente"
    End If
    RsXX.Close
        
End Sub

Private Sub db_FindArticulosFact()
Dim RsXX As rdoResultset
Dim aEnvio As Integer, aNota As Integer, aRemito As Long, aValor As Long, aDevolucion As Long

    On Error GoTo errCargar
    ReDim arrNroSerie(0)
    tArticulo.Clear
    ReDim arrArtCombo(0)
    
    Cons = "Select ArtID, ArtCodigo, ArtBarCode, ArtNombre, ArtNroSerie, ArtTipo, RenCantidad, RenARetirar  From Renglon, Articulo" _
        & " Where RenDocumento = " & Val(lDocumento.Tag) & " And RenArticulo = ArtID "
    Set RsXX = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsXX.EOF
        'Si no es flete lo cargo
        If InStr(Fletes, RsXX!ArtID & ",") = 0 And RsXX!ArtTipo <> paTipoArticuloServicio Then
        
            'Cargo lo que queda en envio, en remito y si hay alguno que este en nota.
            aEnvio = CantidadArticulosEnEnvio(Val(lDocumento.Tag), RsXX!ArtID)
            aRemito = CantidadArticulosEnRemito(Val(lDocumento.Tag), RsXX!ArtID)
            aNota = CantidadArticulosEnNota(Val(lDocumento.Tag), RsXX!ArtID)
            
            aDevolucion = CantidadArticulosEnDevolucion(0, Val(lDocumento.Tag), RsXX!ArtID)
            
            If aDevolucion > 0 Then MsgBox "Ya existe un ingreso de devolución para el artículo " & Trim(RsXX!ArtNombre) & ".", vbInformation, "ATENCIÓN"
        
            If RsXX!RenCantidad > aEnvio + aRemito + aNota + RsXX!RenARetirar + aDevolucion Then
                
                With tArticulo
                    .AddItem Format(RsXX!ArtCodigo, "(#,000,000)") & " " & Trim(RsXX!ArtNombre)
                    .ItemData(.NewIndex) = RsXX!ArtID
                End With
                If arrArtCombo(0).Articulo > 0 Then ReDim Preserve arrArtCombo(UBound(arrArtCombo) + 1)
                With arrArtCombo(UBound(arrArtCombo))
                    .Articulo = RsXX!ArtID
                    .QQuePuede = RsXX!RenCantidad - (aEnvio + aRemito + aNota + RsXX!RenARetirar + aDevolucion)
                    .NroSerie = RsXX!ArtNroSerie
                End With
            End If
        End If
        RsXX.MoveNext
    Loop
    RsXX.Close
    
    If tArticulo.ListCount = 0 Then
        MsgBox "No hay artículos para dar ingreso en este documento." & vbCr & vbCr _
            & " CONSULTE DETALLE DE FACTURA.", vbInformation, "ATENCIÓN"
    End If
    Exit Sub

errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los artículos del documento."
End Sub

Public Function CargoArticulosDeFlete() As String

    On Error GoTo errCargar
    Fletes = ""
    
    'Cargo los articulos a descartar-----------------------------------------------------------
    Cons = "Select Distinct(TFlArticulo) from TipoFlete Where TFlArticulo <> Null"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        Fletes = Fletes & RsAux!TFlArticulo & ","
        RsAux.MoveNext
    Loop
    RsAux.Close
    Fletes = Fletes & paArticuloPisoAgencia & "," & paArticuloDiferenciaEnvio & ","
    '----------------------------------------------------------------------------------------------
    CargoArticulosDeFlete = Fletes
    Exit Function
    
errCargar:
    CargoArticulosDeFlete = Fletes
End Function

Private Function CantidadArticulosEnRemito(lnDocumento As Long, lnArticulo As Long) As Long
On Error GoTo ErrCAER
    'Controlo que la factura no tenga remito x entregar.
    Cons = "Select * From Remito, RenglonRemito Where RemDocumento = " & lnDocumento _
        & " And RReArticulo = " & lnArticulo _
        & " And RemCodigo = RReRemito"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If RsAux.EOF Then
        CantidadArticulosEnRemito = 0
    Else
        CantidadArticulosEnRemito = RsAux!RReAEntregar
    End If
    RsAux.Close
    Exit Function
ErrCAER:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al buscar la cantidad de artículos en remitos."
End Function
Private Function CantidadArticulosEnEnvio(lnDocumento As Long, lnArticulo As Long) As Long
On Error GoTo ErrCAER
    'Controlo que la factura no tenga remito x entregar.
    Cons = "Select Sum(REvAEntregar) From Envio, RenglonEnvio Where EnvDocumento = " & lnDocumento _
        & " And REvArticulo = " & lnArticulo _
        & " And EnvCodigo = REvEnvio"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If IsNull(RsAux(0)) Then
        CantidadArticulosEnEnvio = 0
    Else
        CantidadArticulosEnEnvio = RsAux(0)
    End If
    RsAux.Close
    Exit Function
ErrCAER:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al buscar la cantidad de artículos en envíos."
End Function
Private Function CantidadArticulosEnDevolucion(Cliente As Long, Documento As Long, lnArticulo As Long) As Integer
On Error GoTo ErrCAER
    
    Cons = "Select IsNull(Sum(DevCantidad), 0) From Devolucion  Where DevNota = Null And DevAnulada Is Null And DevArticulo = " & lnArticulo
    
    If Cliente > 0 Then Cons = Cons & " And DevCliente = " & Cliente
    If Documento > 0 Then Cons = Cons & " And DevFactura = " & Documento Else Cons = Cons & " And DevFactura = Null "
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    CantidadArticulosEnDevolucion = RsAux(0)
    RsAux.Close
    Exit Function
ErrCAER:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al buscar la cantidad de artículos en devoluciones."
End Function
Private Function CantidadArticulosEnNota(lnDocumento As Long, lnArticulo As Long) As Long
On Error GoTo ErrCAER
    'Controlo que la factura no tenga remito x entregar.
    Cons = "Select Sum(RenCantidad) From Nota, Documento, Renglon " _
        & " Where NotFactura = " & lnDocumento & " And RenArticulo = " & lnArticulo _
        & " And NotNota = DocCodigo And DocCodigo = RenDocumento And DocAnulado = 0"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If IsNull(RsAux(0)) Then
        CantidadArticulosEnNota = 0
    Else
        CantidadArticulosEnNota = RsAux(0)
    End If
    RsAux.Close
    Exit Function
ErrCAER:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al buscar la cantidad de artículos en remitos."
End Function
Private Sub tUsuario_GotFocus()
    With tUsuario
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tUsuario.Text) Then
            tUsuario.Tag = 0
            tUsuario.Tag = BuscoUsuarioDigito(Val(tUsuario.Text), True)
            If Val(tUsuario.Tag) > 0 And Val(tCantidad.Text) > 0 And tArticulo.ListIndex <> -1 Then AccionGrabar
        Else
            tUsuario.Tag = 0
            MsgBox "Ingrese su dígito de usuario.", vbExclamation, "ATENCIÓN"
        End If
    End If
End Sub

Private Sub PidoNroSerie(Optional bFindProd As Boolean = True)
On Error GoTo errCA
    
    If bFindProd Then
        If HayProductosConNroSerie(Val(lTitular.Tag), tArticulo.ItemData(tArticulo.ListIndex)) Then
            If MsgBox("El cliente posee productos ingresados con número de serie." & vbCr & "¿Desea ingresar el número para validar si esta ingresado?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
                Exit Sub
            End If
        End If
    End If
        
    Dim aNroSerie As String
    Dim intCont As Integer
    For intCont = 1 To Val(tCantidad.Text)
         aNroSerie = ""
        'Input Box Para Nros de Serie
        Do While aNroSerie = ""
            aNroSerie = InputBox("Ingrese el número de serie del artículo entregado." & vbCr & "Con Cancel o vacio abandona el ingreso", "Asociar Producto")
            If Trim(aNroSerie) <> "" Then
                If Not arrAgregoElemento(Trim(aNroSerie)) Then aNroSerie = ""
            Else
                Exit For
            End If
        Loop
    Next
    
    Exit Sub
    
errCA:
    clsGeneral.OcurrioError "Ocurrió un error al insertar el artículo en la grilla.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionGrabar()
Dim bSucCV As Boolean
Dim objSuceso As clsSuceso
Dim lUIDAut As Long, aIdDev As Long
Dim sDefCV As String
    
    If MsgBox("¿Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "GRABAR INGRESO") = vbNo Then Exit Sub
    
    'Cargo los nro de serie.
    PidoNroSerie obIngreso(1).Value Or (obIngreso(0).Value And Not arrArtCombo(tArticulo.ListIndex).NroSerie)
    
    Dim aPendiente As Integer, aCantidad As Integer, aARetirar As Integer
    
    FechaDelServidor
    
    bSucCV = db_CuotasVencidasCliente(IIf(Val(lDocumento.Tag) > 0, Val(lClienteDoc.Tag), Val(lTitular.Tag)), IIf(Val(lDocumento.Tag) > 0, lClienteDoc.Tag, lTitular.Caption), False)
    If bSucCV Then
        Set objSuceso = New clsSuceso
        With objSuceso
            .TipoSuceso = TipoSuceso.DiferenciaDeArticulos
            .ActivoFormulario Val(tUsuario.Tag), "Cliente con Cuotas Atrasadas", cBase
            lUIDAut = .RetornoValor(Usuario:=True)
            If lUIDAut > 0 Then
                tUsuario.Tag = lUIDAut
                sDefCV = .RetornoValor(Defensa:=True)
                If .Autoriza > 0 Then lUIDAut = .Autoriza
            Else
                With tUsuario
                    .Tag = "": .Text = "": .SetFocus
                End With
            End If
        End With
        Set objSuceso = Nothing
        Me.Refresh
        If Val(tUsuario.Tag) = 0 Then Screen.MousePointer = 0: Exit Sub
    End If
    
    'Oculto aquellos que tengan cantidad cero.
    If Val(lDocumento.Tag) > 0 Then
    
        Select Case Val(tCBarra.Tag)
            Case TipoDocumento.Contado, TipoDocumento.Credito
                If CDate(Format(CDate(lFechaDoc.Caption), "dd/mm/yy")) = Date Then
                    If MsgBox("El documento es del día, realmente desea imprimir fichas de devolución por la mercadería ingresada?", vbQuestion + vbYesNo + vbDefaultButton2, "Documento del día") = vbNo Then Exit Sub
                End If
                If CLng(tCantidad.Text) > arrArtCombo(tArticulo.ListIndex).QQuePuede Then
                    MsgBox "La cantidad ingresada es superior a la posible a devolver.", vbExclamation, "Atención"
                    Foco tCantidad
                    Exit Sub
                End If
                
            Case Else
                'Son notas debe coincidir la totalidad.
                If CLng(tCantidad.Text) <> arrArtCombo(tArticulo.ListIndex).QQuePuede Then
                    Screen.MousePointer = 0
                    MsgBox "Esto es una recepción de mercadería por devolución." & _
                        vbCr & "El cliente debe devolver todos los artículos, de lo contrario no podrá realizar el ingreso.", vbExclamation, "Faltan Artículos"
                    Exit Sub
                End If

        End Select
    End If
    
    On Error GoTo ErrBT
    cBase.BeginTrans
    On Error GoTo ErrRB
        
    'Si hay documento, válido que tenga esta cantidad para dar ingreso (no me hayan sacado por otro lado)
    If Val(lDocumento.Tag) > 0 Then
        Cons = "Select * From Documento Where DocCodigo = " & Val(lDocumento.Tag)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux!DocAnulado Then
            RsAux.Close: cBase.RollbackTrans: Screen.MousePointer = 0
            MsgBox "El documento ingresado ha sido anulado. Verifique", vbCritical, "DOCUMENTO ANULADO"
            Exit Sub
        Else
            If Not IsNull(RsAux!DocPendiente) Then
                RsAux.Close: cBase.RollbackTrans: Screen.MousePointer = 0
                MsgBox "La mercadería está pendiente de entrega. Verifique", vbInformation, "ATENCIÓN"
                Exit Sub
            Else
                If RsAux!DocFModificacion = CDate(lFechaDoc.Tag) Then
                    RsAux.Edit
                    RsAux!DocFModificacion = Format(gFechaServidor, sqlFormatoFH)
                    RsAux.Update
                    RsAux.Close
                Else
                    RsAux.Close: cBase.RollbackTrans: Screen.MousePointer = 0
                    MsgBox "El documento fue modificado por otro usuario. Verifique los cambios", vbInformation, "ATENCIÓN"
                    Exit Sub
                End If
            End If
        End If
    End If
        
    If Val(tCBarra.Tag) > 2 And Val(lDocumento.Tag) > 0 Then
        'Pelo todos los artículos que tengo en el array.
        GraboProductosVendidos Val(lDocumento.Tag), True
        GraboDatosTablasDevolucion
        GraboDatosTablaProducto Val(lClienteDoc.Tag)
    Else
        'Pelo todos los artículos que tengo en el array.
        If Val(lDocumento.Tag) > 0 Then GraboProductosVendidos Val(lDocumento.Tag), False
        
        GraboDatosTablaProducto IIf(Val(lDocumento.Tag) > 0, Val(lClienteDoc.Tag), Val(lTitular.Tag))
        
        If Val(lDocumento.Tag) > 0 Then
            Cons = "Select * From Renglon Where RenDocumento = " & Val(lDocumento.Tag) _
                & " And RenArticulo = " & tArticulo.ItemData(tArticulo.ListIndex)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            aCantidad = RsAux!RenCantidad
            aARetirar = RsAux!RenARetirar
            RsAux.Close
            
            aPendiente = CantidadArticulosEnEnvio(Val(lDocumento.Tag), tArticulo.ItemData(tArticulo.ListIndex))
            aPendiente = aPendiente + CantidadArticulosEnRemito(Val(lDocumento.Tag), tArticulo.ItemData(tArticulo.ListIndex))
            aPendiente = aPendiente + CantidadArticulosEnNota(Val(lDocumento.Tag), tArticulo.ItemData(tArticulo.ListIndex))
            aPendiente = aPendiente + CantidadArticulosEnDevolucion(0, Val(lDocumento.Tag), tArticulo.ItemData(tArticulo.ListIndex))
        
            'aCantidad - (aPendiente + aARetirar)
            'Es la cantidad total menos todo lo que tiene por entregar.
            If CLng(tCantidad.Text) > aCantidad - (aPendiente + aARetirar) Then
                cBase.RollbackTrans: Screen.MousePointer = 0
                MsgBox "La cantidad que ingreso es superior a la posible de devolución. Cargue nuevamente el documento.", vbInformation, "ATENCIÓN"
                Exit Sub
            End If
        End If
        
        If Val(lDocumento.Tag) > 0 Then
            Cons = "Select * From Devolucion Where DevFactura = " & Val(lDocumento.Tag) _
                & " And DevNota Is Null And DevArticulo = " & tArticulo.ItemData(tArticulo.ListIndex) _
                & " And DevLocal Is Not Null "
        Else
            Cons = "Select * From Devolucion Where DevCliente = " & Val(lTitular.Tag) _
                & " And DevNota = Null And DevArticulo = " & tArticulo.ItemData(tArticulo.ListIndex) _
                & " And DevLocal Is Not Null And DevFactura Is Null"
        End If
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.AddNew
        RsAux!DevCliente = IIf(Val(lDocumento.Tag) > 0, Val(lClienteDoc.Tag), Val(lTitular.Tag))
        If Val(lDocumento.Tag) > 0 Then RsAux!DevFactura = Val(lDocumento.Tag)
        RsAux!DevArticulo = tArticulo.ItemData(tArticulo.ListIndex)
        RsAux!DevCantidad = tCantidad.Text
        RsAux!DevLocal = paCodigoDeSucursal
        RsAux!DevFAltaLocal = Format(gFechaServidor, sqlFormatoFH)
        If Trim(tComentario.Text) <> "" Then RsAux!DevComentario = Trim(tComentario.Text)
        Cons = GetIDEstados
        If Cons <> "" Then RsAux!DevEstado = Cons
        RsAux.Update
        RsAux.Close
        
        Cons = "Select Max(DevID) From Devolucion Where DevLocal = " & paCodigoDeSucursal _
            & " And DevArticulo = " & tArticulo.ItemData(tArticulo.ListIndex)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not IsNull(RsAux(0)) Then aIdDev = RsAux(0)
        RsAux.Close

        'Hago alta en stock del local y del stock total.
        MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, tArticulo.ItemData(tArticulo.ListIndex), CLng(tCantidad.Text), paEstadoArticuloEntrega, 1
        MarcoMovimientoStockTotal tArticulo.ItemData(tArticulo.ListIndex), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, CLng(tCantidad.Text), 1

        'Si hay documento hago los movimientos con el documento.
        If Val(lDocumento.Tag) > 0 Then
            MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, paCodigoDeSucursal, tArticulo.ItemData(tArticulo.ListIndex), CLng(tCantidad.Text), paEstadoArticuloEntrega, 1, Val(tCBarra.Tag), Val(lDocumento.Tag)
        Else
            MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, paCodigoDeSucursal, tArticulo.ItemData(tArticulo.ListIndex), CLng(tCantidad.Text), paEstadoArticuloEntrega, 1, TipoDocumento.Devolucion, aIdDev
        End If
    End If
    
    
    If bSucCV Then
        clsGeneral.RegistroSucesoAutorizado cBase, gFechaServidor, TipoSuceso.DiferenciaDeArticulos, paCodigoDeTerminal, Val(tUsuario.Tag), Val(lDocumento.Tag), _
                             Descripcion:="Ingreso por Devolución / Cliente debe ctas.", Defensa:=Trim(sDefCV), idCliente:=IIf(Val(lDocumento.Tag) > 0, Val(lClienteDoc.Tag), Val(lTitular.Tag)), idAutoriza:=lUIDAut
    End If
    cBase.CommitTrans
    
    If (Val(tCBarra.Tag) < 3 And Val(lDocumento.Tag) > 0) Or Val(lDocumento.Tag) = 0 Then ImprimoIngresosPorDevolucion aIdDev
    bCancelar_Click
    
    On Error GoTo ErrFin
    
    Exit Sub
ErrBT:
    clsGeneral.OcurrioError "Ocurrió un error al intentar iniciar la transacción.", Err.Description
    Screen.MousePointer = 0
ErrRB:
    Resume ErrVA
ErrVA:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al almacenar la información.", Err.Description
    Exit Sub
ErrFin:
    clsGeneral.OcurrioError "Ocurrió un error al imprimir la ficha y finalizar la grabación de datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub ImprimoIngresosPorDevolucion(ByVal idDevolución As Long)
Dim aTexto As String

    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    
    SeteoImpresoraPorDefecto paPrintConfD
    
    With vsFicha
        .Device = paPrintConfD
        .PaperBin = paPrintConfB
        .PaperSize = paPrintConfPaperSize
        .Orientation = orLandscape
        
        .StartDoc
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión para los retiros." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Sub
        End If
        
        .FileName = "Retiros por Devolucion"
        .FontSize = 8.25
        .TableBorder = tbNone
        
        .FontBold = True
                
        If Val(lDocumento.Tag) > 0 Then
            .TextAlign = taRightBaseline:
            .AddTable ">3000|<3000", "INGRESO DOCUMENTO:|" & lDocumento.Caption, ""
            .TextAlign = taLeftBaseline
        Else
            .Paragraph = " INGRESO SIN DOCUMENTO."
        End If
        
        .Paragraph = ""
        .Paragraph = "Código de Devolución: " & idDevolución
        .Paragraph = "": .Paragraph = ""
        .AddTable "<1800|<4800", "|INGRESO POR DEVOLUCION DE MERCADERIA", ""
         .Paragraph = "": .FontBold = False
         
        .AddTable "<900|<1800", "Fecha:|" & Format(gFechaServidor, "d-Mmm yyyy hh:mm"), ""
        If Val(lDocumento.Tag) = 0 Then
            If tCi.Text <> "" Then
                .AddTable "<900|<5100", "Cliente:|" & Trim(lTitular.Caption) & "(" & clsGeneral.RetornoFormatoCedula(tCi.Text) & ")", ""
            Else
                If tRuc.Text <> "" Then
                    .AddTable "<900|<5100", "Cliente:|" & Trim(lTitular.Caption) & "(" & tRuc.Text & ")", ""
                Else
                    .AddTable "<900|<5100", "Cliente:|" & Trim(lTitular.Caption), ""
                End If
            End If
        End If
        
        .Paragraph = "": .Paragraph = ""
        .Paragraph = "Artículo: " & tArticulo.Text
        .Paragraph = "Cantidad: " & tCantidad.Text
        vsFicha.Paragraph = ""
        If chEstado(0).Visible Then
            
            Dim iQ As Integer
            For iQ = chEstado.LBound To chEstado.UBound
                If aTexto <> "" Then aTexto = aTexto & ",  "
                If chEstado(iQ).Value = 1 Then
                    aTexto = aTexto & Trim(chEstado(iQ).Caption) & " (SI) "
                Else
                    aTexto = aTexto & Trim(chEstado(iQ).Caption) & " (NO) "
                End If
                
                If iQ Mod 2 = 1 Then .Paragraph = aTexto: aTexto = ""
            Next
            If aTexto <> "" Then .Paragraph = aTexto
        End If
        
        vsFicha.Paragraph = ""
        vsFicha.Paragraph = "Comentario:" & tComentario.Text
        
        .EndDoc
        
        .PrintDoc   'Cliente
        .PrintDoc   'Archivo
    End With        '----------------------------------------------------------------------------------------------------------------------------------------------
    
    Screen.MousePointer = 0
    Exit Sub
    
errImprimir:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión de los retiros.", Err.Description
End Sub

Private Sub GraboDatosTablasDevolucion()
Dim aDocumento As Long
Dim rsDev As rdoResultset
    
            
    'Actualizo los datos en tabla Devoluciones------------------------------------------------------------------------------------
    Cons = "Select * From Devolucion" & _
            " Where DevNota = " & Val(lDocumento.Tag) & _
            " And DevArticulo = " & tArticulo.ItemData(tArticulo.ListIndex) & _
            " And DevLocal Is Null"
    Set rsDev = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    rsDev.Edit
    rsDev!DevLocal = paCodigoDeSucursal
    rsDev!DevFAltaLocal = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    Cons = GetIDEstados
    If Cons <> "" Then rsDev!DevEstado = Cons
    rsDev.Update: rsDev.Close
    '-------------------------------------------------------------------------------------------------------------------------------
    
    'Marco el ALTA del STOCK AL LOCAL
    
    'Genero Movimiento
    MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, paCodigoDeSucursal, tArticulo.ItemData(tArticulo.ListIndex), CLng(tCantidad.Text), paEstadoArticuloEntrega, 1, Val(tCBarra.Tag), Val(lDocumento.Tag)
    
    'Alta del Stock en Local
    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, tArticulo.ItemData(tArticulo.ListIndex), CLng(tCantidad.Text), paEstadoArticuloEntrega, 1
    
    'Sumo al Stock Total
    MarcoMovimientoStockTotal tArticulo.ItemData(tArticulo.ListIndex), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, CLng(tCantidad.Text), 1
    
End Sub

Private Function arrAgregoElemento(ByVal aSerie As String) As Boolean
    
    On Error GoTo errAgregar
    arrAgregoElemento = False
    If arrBuscoElemento(aSerie) <> 0 Then
        MsgBox "El nro. de serie ingresado ya fue entregado !!!!." & vbCrLf & vbCrLf & "Nº Serie: " & Trim(aSerie), vbExclamation, "Artículo Entregado"
        Exit Function
    End If
    
    If arrNroSerie(0) <> "" Then ReDim Preserve arrNroSerie(UBound(arrNroSerie) + 1)
    arrNroSerie(UBound(arrNroSerie)) = Trim(aSerie)
    
    arrAgregoElemento = True

errAgregar:
End Function

Private Function arrBuscoElemento(aSerie As String) As Long
    On Error GoTo errB
    arrBuscoElemento = 0
    Dim I As Integer
    For I = LBound(arrNroSerie) To UBound(arrNroSerie)
        If UCase(aSerie) = UCase(arrNroSerie(I)) Then
            arrBuscoElemento = I: Exit Function
        End If
    Next
errB:
End Function

Private Sub GraboDatosTablaProducto(ByVal lCliente As Long)
Dim rsTP As rdoResultset

    'Cambio el cliente en la tabla producto por el cliente Empresa.
    If lCliente = paClienteEmpresa Then Exit Sub
    
    For I = LBound(arrNroSerie) To UBound(arrNroSerie)
        
        If Trim(arrNroSerie(I)) <> "" Then
            Cons = "Select * From Producto Where ProCliente = " & lCliente _
                & " And ProArticulo = " & tArticulo.ItemData(tArticulo.ListIndex) _
                & " And ProNroSerie = '" & Replace(arrNroSerie(I), "'", "''") & "'"
            
            Set rsTP = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rsTP.EOF Then
                rsTP.Edit
                rsTP!ProCliente = paClienteEmpresa
                rsTP!ProFModificacion = Format(Now, "mm/dd/yyyy hh:mm:ss")
                rsTP.Update
            End If
            rsTP.Close
        End If
        
    Next
    
End Sub

Private Sub GraboProductosVendidos(idDocumento As Long, bolEsNota As Boolean, Optional Alta As Boolean = False)
    Dim idFactura As Long
    Dim rsPV As rdoResultset
    
    If Not Alta Then
    
        'Como es una devolucion val(ldocumento.tag) = al IDNOTA --> busco el doc de la nota
        If bolEsNota Then
            Cons = "Select * from Nota Where NotNota = " & idDocumento
            Set rsPV = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rsPV.EOF Then idFactura = rsPV!NotFactura
            rsPV.Close
        Else
            idFactura = idDocumento
        End If
        
        For I = LBound(arrNroSerie) To UBound(arrNroSerie)
            If Trim(arrNroSerie(I)) <> "" Then
                Cons = "Select * from ProductosVendidos Where PVeDocumento = " & idFactura & _
                            " And PVeArticulo = " & tArticulo.ItemData(tArticulo.ListIndex) & _
                            " And PVeNSerie = '" & Replace(Trim(arrNroSerie(I)), "'", "''") & "'"
                Set rsPV = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not rsPV.EOF Then rsPV.Delete
                rsPV.Close
            End If
        Next
    End If
    
End Sub

Private Function HayProductosConNroSerie(ByVal lCliente As Long, ByVal idArt As Long, Optional NroSerie As String = "") As Boolean
On Error GoTo errHPCNS
Dim rsHP As rdoResultset
    HayProductosConNroSerie = False
    Cons = "Select * From Producto Where ProCliente = " & lCliente _
        & " And ProArticulo = " & idArt
    If NroSerie = "" Then
        Cons = Cons & " And ProNroSerie Is Not Null"
    Else
        Cons = Cons & " And ProNroSerie = '" & Trim(NroSerie) & "'"
    End If
    Set rsHP = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsHP.EOF Then HayProductosConNroSerie = True
    rsHP.Close
    Exit Function
errHPCNS:
    clsGeneral.OcurrioError "Ocurrió un error al buscar en la tabla producto.", Trim(Err.Description)
End Function

Private Sub SeteoImpresoraPorDefecto(DeviceName As String)
Dim X As Printer

    For Each X In Printers
        If Trim(X.DeviceName) = Trim(DeviceName) Then
            Set Printer = X
            Exit For
        End If
    Next
    
End Sub

Private Function BuscoUsuarioDigito(Digito As Long, Optional Codigo As Boolean = False, Optional Identificacion As Boolean = False, Optional Iniciales As Boolean = False) As Variant
Dim RsUsr As rdoResultset
Dim aRetorno As Variant
On Error GoTo ErrBUD

    Cons = "Select * from Usuario Where UsuDigito = " & Digito
    Set RsUsr = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsUsr.EOF Then
        If Identificacion Then aRetorno = Trim(RsUsr!UsuIdentificacion)
        If Codigo Then aRetorno = RsUsr!UsuCodigo
        If Iniciales Then aRetorno = Trim(RsUsr!UsuInicial)
    End If
    RsUsr.Close
    BuscoUsuarioDigito = aRetorno
    Exit Function
    
ErrBUD:
    MsgBox "Ocurrio un error inesperado al buscar el usuario.", vbCritical, "ATENCIÓN"
End Function

Private Sub s_SetDocumento()
On Error GoTo errSD
Dim lDoc As Long
    
    tCBarra.Text = UCase(Trim(tCBarra.Text))
    If IsNumeric(Mid(tCBarra.Text, 1, 1)) Then
        'Cod. de barra.
        lDoc = f_FormatoBarras
    Else
        lDoc = db_FindByDocumento
    End If
    If lDoc > 0 Then
        'Cargo los datos del mismo.
        db_LoadDocumento lDoc
    End If
    Exit Sub
errSD:
    clsGeneral.OcurrioError "Error al buscar el documento.", Err.Description
End Sub
Private Sub db_LoadDocumento(ByVal lDoc As Long)
    
    Cons = "Select Documento.*, SucAbreviacion " & _
            " From Documento,  Sucursal" & _
            " Where DocCodigo = " & lDoc & " And DocSucursal = SucCodigo"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        With lDocumento
            .Caption = f_TextoDocumento(RsAux!DocTipo, Trim(RsAux!DocSerie), RsAux!DocNumero)
            If Not IsNull(RsAux!SucAbreviacion) Then .Caption = .Caption & " (" & Trim(RsAux!SucAbreviacion) & ")"
            .Tag = RsAux!DocCodigo
        End With
        
        If RsAux!DocAnulado Then
            MsgBox "El documento está anulado.", vbExclamation, "Atención"
            RsAux.Close
            Exit Sub
        ElseIf Not IsNull(RsAux!DocPendiente) Then
            MsgBox "La mercadería está pendiente de entrega. Verifique", vbInformation, "ATENCIÓN"
            RsAux.Close
            Exit Sub
        Else
            If Not (RsAux!DocTipo <= 4 Or RsAux!DocTipo = 10) Then
                lDocumento.Tag = "": lDocumento.Caption = ""
                MsgBox "Tipo de documento incorrecto.", vbExclamation, "Atención"
                RsAux.Close
                Exit Sub
            End If
        End If
        
        tCBarra.Tag = RsAux!DocTipo
        With lFechaDoc
            .Tag = RsAux!DocFModificacion
            .Caption = Format(RsAux!DocFecha, "d-Mmm-yyyy hh:nn")
        End With
        lClienteDoc.Tag = RsAux!DocCliente
        If Not IsNull(RsAux!DocComentario) Then lComentDoc.Caption = Trim(RsAux!DocComentario)
    Else
        MsgBox "No se encontró el documento ingresado.", vbInformation, "Atención"
        RsAux.Close
        Exit Sub
    End If
    RsAux.Close
    '------------------------------------------------------------------------------
    
    'Cargo cliente
    db_FindCliente Val(lClienteDoc.Tag), lClienteDoc, True
    
    'Artículos del documento.
    If Val(tCBarra.Tag) > 2 Then
        'Busco x nota y ficha de devolución ES EL INGRESO DE LA MERCADERIA HECHA EN FICHA
        db_FindArticuloNota
    Else
        db_FindArticulosFact
    End If
    
    ctrl_SetArticulo tArticulo.ListCount > 0
    If tArticulo.ListCount = 0 Then tCBarra.Tag = ""
    
    
End Sub

Private Function db_FindByDocumento() As Long
On Error GoTo errFD
Dim sSerie As String, sNro As String
    
    If InStr(tCBarra.Text, "-") <> 0 Then
        sSerie = Mid(tCBarra.Text, 1, InStr(tCBarra.Text, "-") - 1)
        sNro = Val(Mid(tCBarra.Text, InStr(tCBarra.Text, "-") + 1))
    Else
        sSerie = Mid(tCBarra.Text, 1, 1)
        sNro = Val(Mid(tCBarra.Text, 2))
    End If
    tCBarra.Text = UCase(sSerie) & "-" & sNro

    db_FindByDocumento = 0
    Cons = "Select DocCodigo, DocFecha as Fecha, DocSerie as Serie, Convert(char(7),DocNumero) as Numero " & _
                " From Documento " & _
                " Where DocTipo IN (1,2,3,4,10)" & _
                " And DocSerie = '" & sSerie & "' And DocNumero = " & sNro & " And DocAnulado = 0"
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        RsAux.MoveNext
        If Not RsAux.EOF Then
            Dim objHelp As New clsListadeAyuda
            With objHelp
                If .ActivarAyuda(cBase, Cons, 5000, 1, "Documentos") > 0 Then
                    db_FindByDocumento = .RetornoDatoSeleccionado(0)
                End If
            End With
            Set objHelp = Nothing
        Else
            RsAux.MoveFirst
            db_FindByDocumento = RsAux(0)
        End If
    Else
        MsgBox "No se encontró un documento con los datos ingresados.", vbInformation, "Atención"
    End If
    RsAux.Close
    Exit Function
errFD:
    clsGeneral.OcurrioError "Error al buscar el documento.", Err.Description
End Function

Private Function f_TextoDocumento(Tipo As Integer, Serie As String, Numero As Long) As String

    Select Case Tipo
        Case 1: f_TextoDocumento = "Ctdo. "
        Case 2: f_TextoDocumento = "Créd. "
        Case 3: f_TextoDocumento = "N/Dev. "
        Case 4: f_TextoDocumento = "N/Créd. "
        Case 5: f_TextoDocumento = "Recibo "
        Case 10: f_TextoDocumento = "N/Esp. "
    End Select
    f_TextoDocumento = f_TextoDocumento & Trim(Serie) & "-" & Numero

End Function

Private Sub db_FindCliente(ByVal Cliente As Long, ByVal ctrlLabel As Label, ByVal bDeDocumento As Boolean)
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

    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    With ctrlLabel
        .Caption = ""
        If Not RsAux.EOF Then
            If Not IsNull(RsAux!CliCIRuc) Then
                If RsAux!CliTipo = TipoCliente.Cliente Then
                    If bDeDocumento Then
                        .Caption = "(" & clsGeneral.RetornoFormatoCedula(RsAux!CliCIRuc) & ") "
                    Else
                        tCi.Text = clsGeneral.RetornoFormatoCedula(RsAux!CliCIRuc)
                        tCi.Tag = 1
                    End If
                Else
                    If bDeDocumento Then
                        .Caption = "(" & clsGeneral.RetornoFormatoRuc(Trim(RsAux!CliCIRuc)) & ") "
                    Else
                        tRuc.Text = Trim(RsAux!CliCIRuc)
                        tRuc.Tag = 2
                    End If
                End If
            ElseIf Not bDeDocumento Then
                If RsAux!CliTipo = TipoCliente.Cliente Then
                    tCi.Tag = 1
                    tRuc.Tag = ""
                Else
                    tCi.Tag = ""
                    tRuc.Tag = "2"
                End If
            End If
            .Tag = Cliente
            .Caption = .Caption & Trim(RsAux!Nombre)
        End If
        RsAux.Close
    End With
    If Not bDeDocumento And Val(ctrlLabel.Tag) > 0 Then ctrl_SetArticulo True
    db_CuotasVencidasCliente Cliente, IIf(bDeDocumento, lClienteDoc.Caption, lTitular.Caption), True
    Exit Sub
    
errCliente:
    clsGeneral.OcurrioError "Error al cargar los datos del cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function db_CuotasVencidasCliente(ByVal lCliente As Long, ByVal sCliente As String, Optional bShowMsg As Boolean) As Boolean
'---------------------------------------------------
'Retorno True si lleva suceso
'---------------------------------------------------
On Error GoTo errCV
Dim rsC As rdoResultset
Dim iMaxAtraso As Integer

    db_CuotasVencidasCliente = False
    
    'Condición para no consultar que el cliente sea de la esta lista.
    If InStr(1, "," & paClienteNoVtoCta & ",", "," & lCliente & ",") > 0 Then
        Exit Function
    End If
    '.......................................................................................
    
    iMaxAtraso = 0
    Cons = "Select Min(CreProximoVto) " & _
                " From Documento (index = iClienteTipo), Credito" & _
                " Where DocCliente = " & lCliente & _
                " And DocCodigo = CreFactura " & _
                " And DocTipo = " & TipoDocumento.Credito & _
                " And DocAnulado = 0  And CreSaldoFactura > 0 "
    
    Set rsC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsC.EOF Then
        If Not IsNull(rsC(0)) Then iMaxAtraso = DateDiff("d", rsC(0), gFechaServidor)
    End If
    rsC.Close
    
    Select Case iMaxAtraso
        Case Is > 20
                If bShowMsg Then MsgBox "El cliente '" & sCliente & "' no está al día." & vbCrLf & _
                            "Tiene coutas vencidas con más de 20 días." & vbCrLf & vbCrLf & _
                            "Consulte antes de realizar el ingreso del artículo.", vbExclamation, "Cliente con Ctas. Vencidas"
                db_CuotasVencidasCliente = True
                
        Case Is > 5
                If bShowMsg Then MsgBox "El cliente '" & sCliente & "' no está al día. Tiene coutas vencidas." & vbCrLf & _
                            "Consulte antes de realizar el ingreso del artículo.", vbExclamation, "Cliente con Ctas. Vencidas"
    End Select
    Exit Function
errCV:
    clsGeneral.OcurrioError "Error al buscar las cuotas vencidas.", Err.Description
End Function

Private Sub ctrl_CleanDocumento()
    tCBarra.Tag = ""
    With lDocumento
        .Caption = ""
        .Tag = ""
    End With
    With lClienteDoc
        .Caption = "": .Tag = ""
    End With
    lComentDoc.Caption = ""
    With lFechaDoc
        .Caption = "": .Tag = ""
    End With
End Sub

Private Sub ctrl_CleanCliente()
    With lTitular
        .Caption = "": .Tag = ""
    End With
    With tRuc
        .Tag = "": .Text = ""
    End With
    With tCi
        .Tag = "": .Text = ""
    End With
    'Al no haber cliente no puede entrar artículos
    ctrl_SetArticulo False
End Sub

Private Sub ctrl_SetArticulo(ByVal bEn As Boolean)
Dim iQ As Integer
    With tArticulo
        .Enabled = bEn: .BackColor = IIf(bEn, vbWhite, vbButtonFace)
    End With
    With tCantidad
        .Enabled = bEn: .BackColor = tArticulo.BackColor
    End With
    If Not bEn Then tArticulo.Clear: tCantidad.Text = ""
    
    If Not chEstado(0).Visible Then Exit Sub
    
    For iQ = chEstado.LBound To chEstado.UBound
        chEstado(iQ).Enabled = bEn
        If Not bEn Then chEstado(iQ).Value = 0
    Next
    
End Sub

Private Sub s_SetCtrlOpt()
    
    With tCBarra
        .Enabled = obIngreso(0).Value
        .BackColor = IIf(obIngreso(0).Value, vbWhite, vbButtonFace)
    End With
    With tCi
        .Enabled = obIngreso(1).Value
        .BackColor = IIf(obIngreso(1).Value, vbWhite, vbButtonFace)
    End With
    With tRuc
        .Enabled = obIngreso(1).Value
        .BackColor = IIf(obIngreso(1).Value, vbWhite, vbButtonFace)
    End With
    
End Sub

Private Function GetIDEstados() As String
Dim iQ As Integer
    
    GetIDEstados = ""
    If chEstado(0).Visible Then
        For iQ = chEstado.LBound To chEstado.UBound
            If GetIDEstados <> "" Then GetIDEstados = GetIDEstados & ":"
            GetIDEstados = GetIDEstados & IIf(chEstado(iQ).Value = 1, "s", "n") & Trim(chEstado(iQ).Tag)
        Next
    End If
    
End Function
