VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{190700F0-8894-461B-B9F5-5E731283F4E1}#1.1#0"; "orHiperlink.ocx"
Begin VB.Form frmEntMercaderia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de mercader�a"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEntMercaderia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VSFlex6DAOCtl.vsFlexGrid lstArticulos 
      Height          =   2895
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5106
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
      ForeColorFixed  =   -2147483640
      BackColorSel    =   13686989
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
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
   Begin VB.TextBox txtArticulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000012&
      Height          =   315
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   17
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox txtNroSerie 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000012&
      Height          =   315
      Left            =   3960
      MaxLength       =   50
      TabIndex        =   16
      Top             =   1320
      Width           =   3495
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7560
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox txtMemo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      MaxLength       =   80
      TabIndex        =   13
      Top             =   4920
      Width           =   5535
   End
   Begin prjHiperLink.orHiperLink hliCliente 
      Height          =   255
      Left            =   960
      Top             =   1080
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   450
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColorOver   =   16711680
      BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjHiperLink.orHiperLink hliDocumento 
      Height          =   255
      Left            =   4200
      Top             =   720
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   450
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColorOver   =   16711680
      MouseIcon       =   "frmEntMercaderia.frx":058A
      MousePointer    =   99
      BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cboEstado 
      Height          =   315
      Left            =   2160
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1440
      Width           =   2775
   End
   Begin VB.PictureBox picTitulo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   8415
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   8445
      Begin VB.ComboBox cboAccion 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   120
         Width           =   4575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Ingreso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.TextBox txtDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000012&
      Height          =   315
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   8415
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5355
      Width           =   8445
      Begin VB.CommandButton butCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   7320
         TabIndex        =   9
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton butAceptar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   6240
         TabIndex        =   8
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblAyuda 
         BackStyle       =   0  'Transparent
         Caption         =   "Devoluci�n de mercader�a ayuda r�pida de lo que tienen que realizar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   450
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   5340
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario"
      Height          =   255
      Left            =   6840
      TabIndex        =   14
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Comentario:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblQueEs 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      Caption         =   " &N�mero de serie"
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
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblDocumento 
      BackStyle       =   0  'Transparent
      Caption         =   "&Documento, C.I./R.U.C."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   750
      Width           =   1935
   End
End
Attribute VB_Name = "frmEntMercaderia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'book vox-076spk1205000323 el de mi nootb.
Option Explicit

Dim sConnect As String

Public Enum Accion
    Informacion = 1         'No toma accion es un comentario +
    Alerta = 2             'Activa la pantalla de comentarios Todas
    Cuota = 3              'Activa en Cobranza, Decision, Visualizacion
    Decision = 4            'Activa en Decision
End Enum


Public Enum Estados
    Anulado = 0
    Visita = 1
    Retiro = 2
    Taller = 3
    Entrega = 4
    Cumplido = 5
End Enum


Public Enum TipoRenglonS
    Llamado = 1
    Cumplido = 2
    CumplidoPresupuesto = 3
    CumplidoArticulo = 4
End Enum

Public Enum TipoAccionEntrada
    TAE_Devolucion = 1
    TAE_Cambio = 2
End Enum

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
   Left     As Long
   Top      As Long
   Right    As Long
   Bottom   As Long
End Type

Dim colArtsDoc As Collection
Dim oEstadoIngMerc As New clsEstadosIngMerc
Dim colRenglonesIngreso As Collection
Dim oArtIngreso As clsRenglonIngreso

Private Function CargoCamposBDRemito(ByVal documento As Long) As Long
Dim mSQL As String

    Dim OBJ_Remito As clsDocumento
    Dim OBJ_Ren  As clsRenglon, colRenglones As New Collection
    
    Set OBJ_Remito = New clsDocumento
    With OBJ_Remito
        .Cliente = Val(hliCliente.Tag)
        .Fecha = gFechaServidor
        .Tipo = 34
        .Sucursal = paCodigoDeSucursal
        .Usuario = Val(txtUser.Tag)
        .Moneda = 1
        
        If Trim(txtMemo.Text) <> "" Then .Comentario = Trim(txtMemo.Text)
    End With
           
    '2) Cargo las estructura de los renglones ---------------------------------------------
    Dim oArt As clsRenglonIngreso
    For Each oArt In colRenglonesIngreso
        
        Set OBJ_Ren = New clsRenglon
        With OBJ_Ren
            .Articulo = oArt.Articulo
            .QCantidad = -1
            .QARetirar = -1
            .QAEnviar = 0
        End With
        colRenglones.Add OBJ_Ren
        Set OBJ_Ren = Nothing
        
    Next
    Set OBJ_Remito.Renglones = colRenglones
    
    Dim objComercio As New clsFunciones
    Set objComercio.Connect = cBase
    
    objComercio.GrabarDocumento paDRemito, OBJ_Remito
    If documento > 0 Then
        objComercio.InsertarRelacionRemitoDocumento documento, OBJ_Remito.Codigo
    End If
    CargoCamposBDRemito = OBJ_Remito.Codigo
    
    Set OBJ_Remito = Nothing
    Set objComercio = Nothing
   
End Function

Private Sub GrabarCambioDeProductosVendidos(ByVal Articulo As clsRenglonIngreso)

    If Trim(Articulo.NroSerieArtCambio) <> "" Then
        'INSERTO EL NUEVO EN LA TABLA PRODUCTOSVENDIDOS
        Cons = "INSERT INTO ProductosVendidos (PVeDocumento, PVeArticulo, PVeNSerie, PVeVarGarantia, PVeVtoGarantia) " & _
                " VALUES(" & Val(hliDocumento.Tag) & ", " & Articulo.Articulo & ", '" & Articulo.NroSerieArtCambio & "', 1, Null)"
        cBase.Execute (Cons)
    End If
    
    If Trim(Articulo.NroSerie) <> "" Then
        'INSERTO o UPDATEO EL VIEJO EN LA TABLA PRODUCTOS VENDIDOS
        Cons = "Select * From ProductosVendidos Where PVeDocumento = " & Val(hliDocumento.Tag) _
                & " And PVeArticulo = " & Articulo.Articulo & " And PVeNSerie = '" & Articulo.NroSerie & "'"
        Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            rsAux.Edit
        Else
            rsAux.AddNew
            rsAux("PVeDocumento") = Val(hliDocumento.Tag)
            rsAux("PVeArticulo") = Articulo.Articulo
            rsAux("PVeNSerie") = Articulo.NroSerie
        End If
        rsAux("PVEVarGarantia") = 255
        rsAux("PVEVtoGarantia") = Now
        rsAux.Update
        rsAux.Close
    End If
    
End Sub

Private Sub InsertoMotivos(ByVal idServicio As Long, ByVal Motivos As Collection)
Dim oMotivo As clsCodigoTexto
    For Each oMotivo In Motivos
        Cons = "Insert Into ServicioRenglon (SReServicio, SReTipoRenglon, SReMotivo, SReCantidad) Values (" & _
            idServicio & ", " & TipoRenglonS.Llamado & ",  " & oMotivo.ID & ", 1)"
        cBase.Execute (Cons)
    Next
End Sub

Private Sub InsertoServicioTaller(idServicio As Long, ByVal LocalRepara As Long)

    If LocalRepara <> paCodigoDeSucursal Then
        'Inserto tambi�n el local para el traslado.
        Cons = "Insert Into Taller(TalServicio, TalFIngresoRealizado, TalFIngresoRecepcion, TalModificacion, TalUsuario, TalLocalAlCliente) Values (" _
            & idServicio & ", GETDATE(), GETDATE()" _
            & ", GETDATE(), " & Val(txtUser.Tag) & ", " & LocalRepara & ")"
    Else
        Cons = "Insert Into Taller(TalServicio, TalFIngresoRealizado, TalFIngresoRecepcion, TalModificacion, TalUsuario) Values (" _
            & idServicio & ", GETDATE(), GETDATE()" _
            & ", GETDATE(), " & Val(txtUser.Tag) & ")"
    End If
    cBase.Execute (Cons)
    
End Sub

Private Function InsertoServicio(ByVal idProducto As Long, ByVal Servicio As clsServicio) As Long
    
    '---------------------------------------------
    'Inserto
    'EstadoP.FueraGarantia = 2
    Cons = "INSERT INTO Servicio (SerProducto, SerFecha, SerEstadoProducto, SerLocalIngreso, " _
        & " SerLocalReparacion, SerEstadoServicio, SerUsuario, SerModificacion, SerCliente, SerComentario) Values (" _
        & idProducto & ", GETDATE(), 2, " & paCodigoDeSucursal
    
    If Servicio.LocalRepara.ID = 0 Then Cons = Cons & ", Null " Else Cons = Cons & ", " & Servicio.LocalRepara.ID
    
    Cons = Cons & ", " & Estados.Taller & ", " & Val(txtUser.Tag) & ", GETDATE(), " & paClienteEmpresa & ", "
    
    If Servicio.Aclaracion = "" Then Cons = Cons & "Null)" Else Cons = Cons & "'" & Servicio.Aclaracion & "')"
    cBase.Execute (Cons)
    
    '---------------------------------------------
    'Saco el mayor c�digo de servicio.
    Cons = "Select Max(SerCodigo) From Servicio"
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    InsertoServicio = rsAux(0)
    rsAux.Close
    '---------------------------------------------
    
End Function


Private Function GraboNuevoProducto(ByVal Articulo As clsRenglonIngreso) As Long
Dim rsTP As rdoResultset
    
    Cons = "Select * From Producto Where ProCodigo = 0"
    Set rsTP = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    rsTP.AddNew
    rsTP!ProArticulo = Articulo.Articulo
    rsTP!ProCliente = paClienteEmpresa
    If Trim(Articulo.NroSerie) <> "" Then rsTP!ProNroSerie = Articulo.NroSerie
    rsTP!ProFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:nn:ss")
    rsTP.Update
    rsTP.Close
    
    'Saco el nuevo ID Para el producto del cliente.
    Cons = "Select Max(ProCodigo) From Producto Where ProCliente = " & paClienteEmpresa & " And ProArticulo = " & Val(Articulo.Articulo)
    Set rsTP = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    GraboNuevoProducto = rsTP(0)
    rsTP.Close

End Function

Private Function CambiarClienteEnTablaProducto(ByVal Articulo As clsRenglonIngreso) As Long
Dim rsTP As rdoResultset
    
    If Val(hliCliente.Tag) = paClienteEmpresa Then Exit Function
    Cons = "Select ProCodigo, ProCliente, ProFModificacion From Producto Where ProCliente = " & Val(hliCliente.Tag) _
        & " And ProArticulo = " & Articulo.Articulo _
        & " And ProNroSerie = '" & Articulo.NroSerie & "'"
    Set rsTP = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsTP.EOF Then
        CambiarClienteEnTablaProducto = rsTP("ProCodigo")
        rsTP.Edit
        rsTP!ProCliente = paClienteEmpresa
        rsTP!ProFModificacion = Format(Now, "mm/dd/yyyy hh:mm:ss")
        rsTP.Update
    End If
    rsTP.Close

End Function

Private Sub InsertoComentarioCambio(ByVal Articulo As String, ByVal Memo As String)
Dim rsM As rdoResultset

    'Le inserto un comentario al documento.
    Cons = "Select * From Comentario Where ComCliente = " & Val(hliCliente.Tag) & " And ComDocumento = " & Val(hliDocumento.Tag) _
        & " And ComUsuario = " & Val(txtUser.Tag)
    
    Set rsM = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    rsM.AddNew
    rsM!ComCliente = Val(hliCliente.Tag)
    rsM!ComFecha = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    rsM!ComComentario = "Producto: " & Trim(Articulo) & Space(5) & "Comentario: " & Trim(Memo)
    rsM!ComTipo = paTipoComentario
    rsM!ComUsuario = Val(txtUser.Tag)
    rsM!ComDocumento = Val(hliDocumento.Tag)
    rsM.Update
    rsM.Close
    
End Sub

Private Sub GrabarIngreso()
Dim Usuario As Integer
Dim autoriza As Integer
Dim defensa As String
    
    On Error GoTo ErrInit
    FechaDelServidor
    If BuscarCuotasVencidasCliente(Val(hliCliente.Tag), hliCliente.Caption, False) Then
        If Not PedirSuceso(Usuario, autoriza, defensa) Then
            Exit Sub
        End If
    End If
    
    On Error GoTo ErrBT
    cBase.BeginTrans
    On Error GoTo ErrRB
    
    Dim TipoDoc As Byte
    Dim documento As Long
    Dim idNewRemito As Long
    
    If Val(hliDocumento.Tag) > 0 Then
        
        'valido fecha de edici�n en el documento.
        Cons = "SELECT DocFModificacion, DocTipo FROM Documento WHERE DocCodigo = " & Val(hliDocumento.Tag) '& _
            " AND DocFModificacion = '" & Format(lblDocumento.Tag, "yyyy/mm/dd HH:nn:ss") & "'"
        Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If rsAux(0) <> CDate(lblDocumento.Tag) Then
            cBase.RollbackTrans
            rsAux.Close
            MsgBox "El documento fu� modificado, no podr� grabar.", vbExclamation, "Atenci�n"
            Exit Sub
        End If
        TipoDoc = rsAux("DocTipo")
        rsAux.Edit
        rsAux("DocFModificacion") = Format(Now, "yyyy/mm/dd HH:mm:ss")
        rsAux.Update
        rsAux.Close

        If TipoDoc = 6 Then
            Cons = "SELECT RDoDocumento FROM RemitoDocumento WHERE RDoRemito = " & Val(hliDocumento.Tag)
            Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rsAux.EOF Then documento = rsAux(0) Else documento = hliDocumento.Tag
            rsAux.Close
        Else
            documento = hliDocumento.Tag
        End If
        
    End If
    
    'Si es cambio de producto me quedo con el documento.
'    documento = Val(hliDocumento.Tag)

    Dim oArtAFicha As New Collection
    
    Dim oArt As clsRenglonIngreso
    For Each oArt In colRenglonesIngreso
    
        Dim idProducto As Long
        idProducto = CambiarClienteEnTablaProducto(oArt)
        If oArt.Servicio.LocalRepara.ID > 0 Then
            'Inserto servicio
            '1) Si no tengo producto lo inserto.
            If idProducto = 0 Then idProducto = GraboNuevoProducto(oArt)
            Dim idServicio As Long
            idServicio = InsertoServicio(idProducto, oArt.Servicio)
            If oArt.Servicio.Motivos.Count > 0 Then InsertoMotivos idServicio, oArt.Servicio.Motivos
            If oArt.Servicio.LocalRepara.ID = paCodigoDeSucursal Then InsertoServicioTaller idServicio, oArt.Servicio.LocalRepara.ID
            oArt.Servicio.IDNuevoServicio = idServicio
        End If
        
        If cboAccion.ItemData(cboAccion.ListIndex) = TAE_Cambio Then
            'CAMBIO DE PRODUCTO
            'Tengo que hacer la salida del art�culo.
            Cons = "EXEC stockMovFisicoEnLocal " & Val(txtUser.Tag) & ", " & oArt.Articulo & ", -1, " & paEstadoArticuloEntrega & ", 2, " & paCodigoDeSucursal & _
                    ", " & 27 & ", " & documento & ", " & paCodigoDeTerminal
            cBase.Execute Cons

            Cons = "EXEC StockMovEstadoStockTotal " & Val(txtUser.Tag) & ", " & oArt.Articulo & ", -1, " & paEstadoArticuloEntrega _
                & ", 1, 0"
            cBase.Execute Cons

            'Si tengo documento grabo tabla productos vendidos (sin doc es s�lo si se hace por CI y no por DOC).
            If Val(hliDocumento.Tag) > 0 Then GrabarCambioDeProductosVendidos oArt
            InsertoComentarioCambio oArt.ArticuloNombre, IIf(txtMemo.Text <> "", txtMemo.Text, oArt.Servicio.Aclaracion)
            
        Else
            
            If TipoDoc <= 2 Or TipoDoc = 6 Then
                'Recorro la lista para ver si ya tengo otro.
                Dim oArtF As clsArticuloAFicha
                Dim oArtN As clsArticuloAFicha
                Set oArtN = Nothing
                For Each oArtF In oArtAFicha
                    If oArtF.Articulo = oArt.Articulo Then
                        Set oArtN = oArtF
                        Exit For
                    End If
                Next
                If oArtN Is Nothing Then
                    Set oArtN = New clsArticuloAFicha
                    oArtAFicha.Add oArtN
                    oArtN.Articulo = oArt.Articulo
                End If
                oArtN.Cantidad = oArtN.Cantidad + 1
                oArtN.EstadoMensaje = oEstadoIngMerc.ObtenerStringEstado(oArt.Estados)

'                Dim aIDDev As Long
'                If documento > 0 Then
'                    Cons = "Select * From Devolucion Where DevFactura = " & documento _
'                        & " And DevNota Is Null And DevArticulo = " & oArt.Articulo _
'                        & " And DevLocal Is Not Null "
'                Else
'                    Cons = "Select * From Devolucion Where DevCliente = " & Val(hliCliente.Tag) _
'                        & " And DevNota = Null And DevArticulo = " & oArt.Articulo _
'                        & " And DevLocal Is Not Null And DevFactura Is Null"
'                End If
'                Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'                rsAux.AddNew
'                rsAux!DevCliente = Val(hliCliente.Tag)
'                If documento > 0 Then rsAux!DevFactura = documento
'                rsAux!DevArticulo = oArt.Articulo
'                rsAux!DevCantidad = 1
'                rsAux!DevLocal = paCodigoDeSucursal
'                rsAux!DevFAltaLocal = Format(gFechaServidor, "yyyy/MM/dd HH:nn:ss")
'                If Trim(txtMemo.Text) <> "" Then rsAux!DevComentario = Trim(txtMemo.Text)
'                rsAux!DevEstado = oEstadoIngMerc.ObtenerStringEstado(oArt.Estados) 'oEstadoIngMerc.ObtenerCadenaEstadosSeleccionados(oArt.estado)
'                rsAux.Update
'                rsAux.Close
'
'                Cons = "Select Max(DevID) From Devolucion Where DevLocal = " & paCodigoDeSucursal _
'                    & " And DevArticulo = " & oArt.Articulo
'                Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'                If Not IsNull(rsAux(0)) Then aIDDev = rsAux(0)
'                rsAux.Close
                
'                oArt.FichaDevolucion = aIDDev
            End If
            
        End If
        
    'SI ES ENTRADA O CAMBIO DE PRODUCTO EL ESTADO DEPENDE SI LE INDICAN REPARAR.
    '---------------------------------------------------------
    
'        If bHacerMovFisico Then
            'Hago el ingreso al local del art�culo que ingresa.
            Cons = "EXEC stockMovFisicoEnLocal " & Val(txtUser.Tag) & ", " & oArt.Articulo & ", 1, " & _
                    IIf(oArt.Servicio.LocalRepara.ID > 0, paEstadoARecuperar, IIf(oArt.EsRoto, paEstadoRoto, paEstadoArticuloEntrega)) & _
                    ", 2, " & paCodigoDeSucursal & _
                    ", " & 27 & ", " & documento & ", " & paCodigoDeTerminal
                    '", " & IIf(cboAccion.ItemData(cboAccion.ListIndex) = TAE_Cambio, 27, 28) & ", " & IIf(cboAccion.ItemData(cboAccion.ListIndex) = TAE_Cambio, documento, oArt.FichaDevolucion) & ", " & paCodigoDeTerminal
            cBase.Execute Cons
            
            'StockMovEstadoStockTotal @iUser smallint, @iArticulo int,  @iCantidad smallint = 0, @iEstado smallint, @iTEst tinyint, @iTQ tinyint = 0
            Cons = "EXEC StockMovEstadoStockTotal " & Val(txtUser.Tag) & ", " & oArt.Articulo & ", 1, " & _
                    IIf(oArt.Servicio.LocalRepara.ID > 0, paEstadoARecuperar, IIf(oArt.EsRoto, paEstadoRoto, paEstadoArticuloEntrega)) _
                & ", 1, 0"
            cBase.Execute Cons
'        End If
    '---------------------------------------------------------
        
    Next
    
    Dim oArtFicha As clsArticuloAFicha
    For Each oArtFicha In oArtAFicha
        Dim aIDDev As Long
        If documento > 0 Then
            Cons = "Select * From Devolucion Where DevFactura = " & documento _
                & " And DevNota Is Null And DevArticulo = " & oArtFicha.Articulo _
                & " And DevLocal Is Not Null "
        Else
            Cons = "Select * From Devolucion Where DevCliente = " & Val(hliCliente.Tag) _
                & " And DevNota = Null And DevArticulo = " & oArtFicha.Articulo _
                & " And DevLocal Is Not Null And DevFactura Is Null"
        End If
        Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        rsAux.AddNew
        rsAux!DevCliente = Val(hliCliente.Tag)
        If documento > 0 Then rsAux!DevFactura = documento
        rsAux!DevArticulo = oArtFicha.Articulo
        rsAux!DevCantidad = oArtFicha.Cantidad
        rsAux!DevLocal = paCodigoDeSucursal
        rsAux!DevFAltaLocal = Format(gFechaServidor, "yyyy/MM/dd HH:nn:ss")
        
        If Trim(txtMemo.Text) <> "" Then rsAux!DevComentario = Trim(txtMemo.Text)
        rsAux!DevEstado = oArtFicha.EstadoMensaje
        rsAux.Update
        rsAux.Close

        Cons = "Select Max(DevID) From Devolucion Where DevLocal = " & paCodigoDeSucursal _
            & " And DevArticulo = " & oArtFicha.Articulo & " AND DevCantidad = " & oArtFicha.Cantidad
        Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        oArtFicha.IDFicha = rsAux(0)
        rsAux.Close
    Next
    If Usuario > 0 Then
        clsGeneral.RegistroSucesoAutorizado cBase, gFechaServidor, 11, paCodigoDeTerminal, Usuario, Val(hliDocumento.Tag), _
                             Descripcion:="Ingreso " & Trim(cboAccion.Text) & " / Cliente debe ctas.", defensa:=Trim(defensa), idCliente:=Val(hliCliente.Tag), idAutoriza:=CLng(autoriza)
    End If
    cBase.CommitTrans
    
    On Error GoTo errFin
    ImprimirFichas oArtAFicha
    
    Set oArtAFicha = Nothing
    On Error Resume Next
    butCancelar_Click
    
Exit Sub
errFin:
    clsGeneral.OcurrioError "Error al imprimir o al ajustar el formulario.", Err.Description, "Grabar"
    Screen.MousePointer = 0
    Exit Sub
    
ErrInit:
    clsGeneral.OcurrioError "Error al intentar grabar la informaci�n.", Err.Description, "Grabar"
    Screen.MousePointer = 0
    Exit Sub
ErrBT:
    clsGeneral.OcurrioError "Ocurri� un error al intentar iniciar la transacci�n.", Err.Description
    Screen.MousePointer = 0
ErrRB:
    Resume ErrVA
ErrVA:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurri� un error al almacenar la informaci�n.", Err.Description
    Exit Sub
End Sub

Private Sub ImprimirFichas(ByVal ArtsFichaDev As Collection)
On Error GoTo errFD
Dim iCont As Integer
Dim oPrint As clsPrintReport
    Set oPrint = New clsPrintReport
    With oPrint
        .StringConnect = sConnect
        .DondeImprimo.Bandeja = paPrintConfB
        .DondeImprimo.Impresora = paPrintConfD
        .DondeImprimo.Papel = paPrintConfPaperSize
        .PathReportes = paPathReportes
    End With
    
    Dim sQueryFicha As String
    sQueryFicha = "SELECT 'INGRESO ' + CASE WHEN DevFactura IS NULL THEN 'SIN DOCUMENTO' ELSE ' CON DOCUMENTO ' + RTRIM(DocSerie) + ' ' + CONVERT(varchar(6), DocNumero) END infoTitulo " & _
        ", ArtCodigo infoCodigoArt, RTRIM(ArtNombre) infoNombreArt, DevCantidad infoCantidad, dbo.FormatDate(DevFAltaLocal, 121) infoFecha " & _
        ", CASE WHEN DevFactura IS NULL THEN '' ELSE 'Documento ' + RTRIM(DocSerie) + ' ' + CONVERT(varchar(6), DocNumero) END infoDocumento " & _
        ", 'Devoluci�n: ' + CONVERT(varchar(8), DevID) infoCodigoDeFicha, IsNull(DevComentario, '') infoComentario " & _
        ", '{1}' infoEstado " & _
        "FROM CGSA.dbo.Devolucion LEFT OUTER JOIN CGSA.dbo.Documento ON DevFactura = DocCodigo " & _
        "INNER JOIN CGSA.dbo.Articulo ON DevArticulo = ArtId WHERE DevID = {0}"

    Dim sQueryServicio As String
    sQueryServicio = "SELECT SerCodigo infoCodigoServicio, '*S'+ RTRIM(convert(varchar(20), SerCodigo)) + '*' infoCodigoBarras " & _
            ", dbo.FormatCIRuc(CliCIRUC) + ' ' + CASE WHEN CliTipo = 1 THEN rtrim(CPeNombre1) + ' ' + RTRIM(CPeApellido1) ELSE RTRIM(CEmNombre) END infoCliente " & _
            ", dbo.FormatDate(SerFecha, 121) infoFecha " & _
            ",'(' + rtrim(Convert(varchar(10), ProCodigo)) + ') ' + RTRIM(ArtNombre) + ' (' + Convert(varchar(15), ArtCodigo) + ')' infoArticulo " & _
            ", RTRIM(usuidentificacion) infoRecibio " & _
            ", dbo.TelefonosCliente(CliCodigo) infoTelefonos, '{1}' infoMotivos, IsNull(SerComentario, '') infoMemoIngreso " & _
            ", dbo.ArmoDireccion(CliDireccion) infoDireccion, IsNull(IsNull(ProFacturaS, '') + ' ' + CONVERT(varchar(6), ProFacturaN), '') infoFactura " & _
            ", Rtrim(SucI.SucAbreviacion) infoLocal, Rtrim(IsNull(SucS.SucAbreviacion, '')) infoLocalRepara, RTRIM(ProNroSerie) infoNroSerie, ISNull(dbo.FormatDate(ProCompra, 2), '') infoFCompra " & _
            "FROM Servicio INNER JOIN Cliente ON SerCliente = CliCodigo " & _
            "LEFT OUTER JOIN CPersona ON CliCodigo = CPeCliente " & _
            "LEFT OUTER JOIN CEmpresa ON CliCodigo = CEmCliente " & _
            "INNER JOIN Producto ON SerProducto = ProCodigo " & _
            "INNER JOIN Articulo ON ProArticulo = ArtId " & _
            "INNER JOIN Sucursal SucI ON SerLocalIngreso = SucI.SucCodigo " & _
            "LEFT OUTER JOIN Sucursal SucS ON SerLocalReparacion = SucS.SucCodigo " & _
            "INNER JOIN Usuario ON SerUsuario = UsuCodigo " & _
            "WHERE SerCodigo = {0}"
            
    Dim sQueryServicioSinNro As String
    sQueryServicioSinNro = "SELECT ' sin n�mero' infoCodigoServicio, '' infoCodigoBarras " & _
            ", '" & hliCliente.Caption & "' infoCliente " & _
            ", dbo.FormatDate(GETDATE(), 121) infoFecha " & _
            ", '{0}' infoArticulo , '{1}' infoMotivos, '{2}' infoNroSerie " & _
            ", RTRIM(UsuIdentificacion) infoRecibio " & _
            ", dbo.TelefonosCliente(" & Val(hliCliente.Tag) & ") infoTelefonos " & _
            ", dbo.ArmoDireccion(" & Val(hliCliente.Tag) & ") infoDireccion " & _
            ", '" & paNombreSucursal & "' infoLocal " & _
            "FROM Usuario WHERE UsuCodigo = " & Val(txtUser.Tag)
    
    Dim query As String
    Dim oArtL As clsRenglonIngreso
    For Each oArtL In colRenglonesIngreso
'        If oArtL.FichaDevolucion > 0 Then
'            query = Replace(sQueryFicha, "{0}", oArtL.FichaDevolucion)
'            query = Replace(query, "{1}", oEstadoIngMerc.ObtenerCadenaEstadosSeleccionados(oArtL.Estados))
'            oPrint.Imprimir_vsReport "FichaDevolucion.xml", "FichaDeDevolucion", query, "", ""
'            oPrint.Imprimir_vsReport "FichaDevolucion.xml", "FichaDeDevolucion", query, "", ""
'        End If
        If oArtL.Servicio.Vias > 0 Then
            If oArtL.Servicio.IDNuevoServicio > 0 Then
                query = Replace(sQueryServicio, "{0}", oArtL.Servicio.IDNuevoServicio)
                query = Replace(query, "{1}", MotivosCargados(oArtL.Servicio.Motivos)) '   MotivosDelServicio(oArtL.Servicio.IDNuevoServicio))
            Else
                query = Replace(sQueryServicioSinNro, "{0}", oArtL.ArticuloNombre)
                query = Replace(query, "{1}", MotivosCargados(oArtL.Servicio.Motivos))
                query = Replace(query, "{2}", oArtL.NroSerie)
            End If
            oPrint.Imprimir_vsReport "FichaServicio.xml", "FichaDeServicio", query, "", ""
            If oArtL.Servicio.Vias = 2 Then
                oPrint.Imprimir_vsReport "FichaServicio.xml", "FichaDeServicio", query, "", ""
            End If
        End If
    Next
    
    Dim oDev As clsArticuloAFicha
    For Each oDev In ArtsFichaDev
        If oDev.IDFicha > 0 Then
            query = Replace(sQueryFicha, "{0}", oDev.IDFicha)
            If (oDev.EstadoMensaje = "") Then
                query = Replace(query, "{1}", "Inmaculado")
            Else
                query = Replace(query, "{1}", oEstadoIngMerc.ObtenerCadenaEstadosSeleccionados(oDev.EstadoMensaje))
            End If
            oPrint.Imprimir_vsReport "FichaDevolucion.xml", "FichaDeDevolucion", query, "", ""
            oPrint.Imprimir_vsReport "FichaDevolucion.xml", "FichaDeDevolucion", query, "", ""
        End If
    Next
    Exit Sub
    
errFD:
    clsGeneral.OcurrioError "Error al imprimir las fichas.", Err.Description, "Fichas de devoluci�n"
End Sub

Private Function MotivosCargados(ByVal Motivos As Collection) As String
    Dim Motivo As clsCodigoTexto
    For Each Motivo In Motivos
        MotivosCargados = MotivosCargados & IIf(MotivosCargados <> "", ", ", "") & Motivo.Nombre
    Next
End Function


Private Function MotivosDelServicio(ByVal idServicio As Long) As String
Dim rsM As rdoResultset
Dim query As String
    query = "SELECT RTRIM(MSeNombre) FROM ServicioRenglon INNER JOIN MotivoServicio ON SReMotivo = MSeID WHERE SReServicio = " & idServicio
    Set rsM = cBase.OpenResultset(query, rdOpenDynamic, rdConcurValues)
    Do While Not rsM.EOF
        MotivosDelServicio = MotivosDelServicio & IIf(MotivosDelServicio <> "", ", ", "") & rsM(0)
        rsM.MoveNext
    Loop
    rsM.Close
End Function

'Private Sub GrabarIngresoConRemito()
'Dim Usuario As Integer
'Dim autoriza As Integer
'Dim defensa As String
'
'    On Error GoTo ErrInit
'    FechaDelServidor
'    If BuscarCuotasVencidasCliente(Val(hliCliente.Tag), hliCliente.Caption, False) Then
'        If Not PedirSuceso(Usuario, autoriza, defensa) Then
'            Exit Sub
'        End If
'    End If
'
'    On Error GoTo ErrBT
'    cBase.BeginTrans
'    On Error GoTo ErrRB
'
'    Dim TipoDoc As Byte
'    Dim documento As Long
'    Dim idNewRemito As Long
'
'    If Val(hliDocumento.Tag) > 0 Then
'
'        'valido fecha de edici�n en el documento.
'        Cons = "SELECT DocFModificacion, DocTipo FROM Documento WHERE DocCodigo = " & Val(hliDocumento.Tag) & _
'            " AND DocFModificacion = '" & Format(lblDocumento.Tag, "yyyy/mm/dd HH:nn:ss") & "'"
'        Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'        If rsAux(0) <> CDate(lblDocumento.Tag) Then
'            cBase.RollbackTrans
'            rsAux.Close
'            MsgBox "El documento fu� modificado, no podr� grabar.", vbExclamation, "Atenci�n"
'            Exit Sub
'        End If
'        TipoDoc = rsAux("DocTipo")
'        rsAux.Edit
'        rsAux("DocFModificacion") = Format(Now, "yyyy/mm/dd HH:mm:ss")
'        rsAux.Update
'        rsAux.Close
'
'        If TipoDoc > 2 Then
'            'Si es un remito --> busco el id del documento asociado.
'
'            'Busco el id de la factura.
'            Cons = "SELECT RDoDocumento FROM RemitoDocumento WHERE RDoRemito = " & Val(hliDocumento.Tag)
'            Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'            If Not rsAux.EOF Then documento = rsAux(0)
'            rsAux.Close
'
'        Else
'
'            documento = hliDocumento.Tag
'
'        End If
'    End If
'
'    'Si es devoluci�n de art�culos tengo que generar el remito.
'    If cboAccion.ItemData(cboAccion.ListIndex) = 1 Then
'
'        'Ojo puede ser ingreso por documento o por cliente.
'        'Por lo que si tiene documento se asocia si no queda libre.
'        If TipoDoc <> 34 Then
'            idNewRemito = CargoCamposBDRemito(documento)
'        Else
'            idNewRemito = Val(hliDocumento.Tag)
'        End If
'
'    Else
'
'        'Si es cambio de producto me quedo con el documento.
'        documento = Val(hliDocumento.Tag)
'
'    End If
'
'
'
'    Dim oArt As clsRenglonIngreso
'    For Each oArt In colRenglonesIngreso
'
''        ALTER  PROCEDURE stockMovFisicoEnLocal @iUser smallint, @iArticulo int, @iCantidad smallint, @iEstado smallint, @iTipoLocal tinyint, @iLocal int,
''        @iTipoDocumento smallint = Null,  @iDocumento Int = Null, @iTerminal smallint = Null
'
'
'        'SI ES ENTRADA O CAMBIO DE PRODUCTO EL ESTADO DEPENDE SI LE INDICAN REPARAR.
'    '---------------------------------------------------------
'
'        'Hago el ingreso al local del art�culo que ingresa.
'        Cons = "EXEC stockMovFisicoEnLocal " & Val(txtUser.Tag) & ", " & oArt.Articulo & ", 1, " & IIf(oArt.Servicio.LocalRepara.ID > 0, paEstadoARecuperar, paEstadoArticuloEntrega) & _
'                ", 2, " & paCodigoDeSucursal & _
'                ", " & IIf(idNewRemito > 0, 34, 0) & ", " & idNewRemito & ", " & paCodigoDeTerminal
'        cBase.Execute Cons
'
'        'StockMovEstadoStockTotal @iUser smallint, @iArticulo int,  @iCantidad smallint = 0, @iEstado smallint, @iTEst tinyint, @iTQ tinyint = 0
'        Cons = "EXEC StockMovEstadoStockTotal " & Val(txtUser.Tag) & ", " & oArt.Articulo & ", 1, " & IIf(oArt.Servicio.LocalRepara.ID > 0, paEstadoARecuperar, paEstadoArticuloEntrega) _
'            & ", 1, 0"
'        cBase.Execute Cons
'    '---------------------------------------------------------
'
'        If cboAccion.ItemData(cboAccion.ListIndex) = 2 Then
'            'CAMBIO DE PRODUCTO
'            'Tengo que hacer la salida del art�culo.
'            Cons = "EXEC stockMovFisicoEnLocal " & Val(txtUser.Tag) & ", " & oArt.Articulo & ", -1, " & paEstadoArticuloEntrega & ", 2, " & paCodigoDeSucursal & _
'                    ", " & TipoDoc & ", " & documento & ", " & paCodigoDeTerminal & ")"
'
'
'            Cons = "EXEC StockMovEstadoStockTotal " & Val(txtUser.Tag) & ", " & oArt.Articulo & ", -1, " & paEstadoArticuloEntrega _
'                & ", 1, 0"
'
''            If oArt.NroSerieArtCambio <> "" Then
'    'tengo que ver que hago con esto cuando entrega la mercaderia
''                GrabarCambioDeProductosVendidos oArt
''            End If
'
'        End If
'
'
'    Next
'
'    If Usuario > 0 Then
'        clsGeneral.RegistroSucesoAutorizado cBase, gFechaServidor, 11, paCodigoDeTerminal, Usuario, Val(hliDocumento.Tag), _
'                             Descripcion:="Ingreso " & Trim(cboAccion.Text) & " / Cliente debe ctas.", defensa:=Trim(defensa), idCliente:=Val(hliCliente.Tag), idAutoriza:=CLng(autoriza)
'    End If
'    cBase.CommitTrans
'
'    On Error Resume Next
'    butCancelar_Click
'
'Exit Sub
'ErrInit:
'    clsGeneral.OcurrioError "Error al intentar grabar la informaci�n.", Err.Description, "Grabar"
'    Screen.MousePointer = 0
'    Exit Sub
'ErrBT:
'    clsGeneral.OcurrioError "Ocurri� un error al intentar iniciar la transacci�n.", Err.Description
'    Screen.MousePointer = 0
'ErrRB:
'    Resume ErrVA
'ErrVA:
'    cBase.RollbackTrans
'    Screen.MousePointer = 0
'    clsGeneral.OcurrioError "Ocurri� un error al almacenar la informaci�n.", Err.Description
'    Exit Sub
'End Sub


Private Function PedirSuceso(ByRef Usuario As Integer, ByRef autoriza As Integer, ByRef defensa As String) As Boolean
    
    Usuario = 0
    
    Dim objSuceso As New clsSuceso
    With objSuceso
        .TipoSuceso = 11 ' TipoSuceso.DiferenciaDeArticulos
        .ActivoFormulario Val(txtUser.Tag), "Cliente con Cuotas Atrasadas", cBase
        Usuario = .RetornoValor(Usuario:=True)
        If Usuario > 0 Then
            defensa = .RetornoValor(defensa:=True)
            If .autoriza > 0 Then autoriza = .autoriza
        End If
    End With
    Set objSuceso = Nothing
    Me.Refresh
    PedirSuceso = (Usuario > 0)

End Function

Private Function BuscarCuotasVencidasCliente(ByVal lCliente As Long, ByVal sCliente As String, Optional bShowMsg As Boolean) As Boolean
'---------------------------------------------------
'Retorno True si lleva suceso
'---------------------------------------------------
On Error GoTo errCV
Dim rsC As rdoResultset
Dim iMaxAtraso As Integer

    BuscarCuotasVencidasCliente = False
    
    'Condici�n para no consultar que el cliente sea de la esta lista.
    If InStr(1, "," & paClienteNoVtoCta & ",", "," & lCliente & ",") > 0 Then
        Exit Function
    End If
    '.......................................................................................
    
    iMaxAtraso = 0
    Cons = "Select Min(CreProximoVto) " & _
                " From Documento (index = iClienteTipo), Credito" & _
                " Where DocCliente = " & lCliente & _
                " And DocCodigo = CreFactura " & _
                " And DocTipo = 2" & _
                " And DocAnulado = 0  And CreSaldoFactura > 0 "
    
    Set rsC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsC.EOF Then
        If Not IsNull(rsC(0)) Then iMaxAtraso = DateDiff("d", rsC(0), Now)
    End If
    rsC.Close
    
    Select Case iMaxAtraso
        Case Is > 20
                If bShowMsg Then MsgBox "El cliente '" & sCliente & "' no est� al d�a." & vbCrLf & _
                            "Tiene coutas vencidas con m�s de 20 d�as." & vbCrLf & vbCrLf & _
                            "Consulte antes de realizar el ingreso del art�culo.", vbExclamation, "Cliente con Ctas. Vencidas"
                BuscarCuotasVencidasCliente = True
                
        Case Is > 5
                If bShowMsg Then MsgBox "El cliente '" & sCliente & "' no est� al d�a. Tiene coutas vencidas." & vbCrLf & _
                            "Consulte antes de realizar el ingreso del art�culo.", vbExclamation, "Cliente con Ctas. Vencidas"
    End Select
    Exit Function
    
errCV:
    clsGeneral.OcurrioError "Error al buscar las cuotas vencidas.", Err.Description
End Function

Private Function ExisteNroSerieEnGrilla(ByVal serie As String) As Boolean
    Dim oArt As clsRenglonIngreso
    If serie = "0" Then Exit Function
    For Each oArt In colRenglonesIngreso
        If oArt.NroSerie = serie Or oArt.NroSerie = serie And oArtIngreso.Articulo = oArt.Articulo Then
            ExisteNroSerieEnGrilla = True
            Exit Function
        End If
    Next
End Function

Private Sub CargarLocalesRepara()
On Error GoTo errCLR
    Screen.MousePointer = 11
    Dim oLocal As clsCodigoTexto
    Set oLocal = New clsCodigoTexto
    oLocal.Nombre = "No Reparar"
    oLocal.ID = 0
    colLocalesRepara.Add oLocal
    
    Cons = "SELECT IsNull(rTRIM(ParTexto), '') FROM Parametro WHERE ParNombre = 'LocalesDeService'"
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    If Not rsAux.EOF Then
        Cons = rsAux(0)
    Else
        Cons = ""
    End If
    rsAux.Close
    
    'Agrego local compa�ia.
    If Cons <> "" Then Cons = Cons & ","
    Cons = Cons & paLocalCompa�ia & ", " & paCodigoDeSucursal
    
    Cons = " SELECT SucCodigo, RTrim(SucAbreviacion)" & _
        " From Sucursal WHERE SucCodigo IN (" & Cons & ")" & _
        " Order by SucAbreviacion"
    
    
'    cons = "DECLARE @txt varchar(100) SELECT @txt=rTRIM(ParTexto) FROM Parametro WHERE ParNombre = 'LocalesDeService' "
'    cons = cons & " SELECT SucCodigo, RTrim(SucAbreviacion)" & _
'        " From Sucursal INNER JOIN dbo.InTable(@txt) ON SucCodigo = convert(smallint, valor) Order by SucAbreviacion"
''        WHERE SucCodigo IN (SELECT rTRIM(ParTexto) FROM Parametro WHERE ParNombre = 'LocalesDeService')" & _
''        " Order by SucAbreviacion"
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    Do While Not rsAux.EOF
        Set oLocal = New clsCodigoTexto
        oLocal.Nombre = rsAux(1)
        oLocal.ID = rsAux(0)
        colLocalesRepara.Add oLocal
        rsAux.MoveNext
    Loop
    rsAux.Close
    Screen.MousePointer = 0
    Exit Sub
errCLR:
clsGeneral.OcurrioError "Error al cargar los locales de reparaci�n.", Err.Description, "Cargar Locales"
Screen.MousePointer = 0
End Sub
Private Function BuscarArticulo(ByVal showmsg As Boolean, ByVal filter As String) As Boolean
On Error GoTo errBA

Dim rsArt As rdoResultset

    Set oArtIngreso = New clsRenglonIngreso
    
    Cons = "EXEC prg_BuscarArticuloEscaneado '" & filter & "'"
    
    Set rsArt = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not IsNull(rsArt("ArtID")) Then
        oArtIngreso.Articulo = rsArt("ArtID")
        oArtIngreso.ArticuloNombre = Trim(rsArt("ArtNombre"))
        
        'Veo si es el nro. de serie del art�culo.
        If rsArt("ACBLargo") > 0 And Val(rsArt("ArtCodigo")) <> Val(filter) Then
            oArtIngreso.NroSerie = filter
            oArtIngreso.TipoNroSerie = 1
        ElseIf rsArt("PedirNSerie") Then
            oArtIngreso.TipoNroSerie = 2
        End If
        
        'Este SP me puede retonar un art. espec�fico.
        If rsArt("AEsID") > 0 Then
            
            Cons = "SELECT RTrim(IsNull(AEsNroSerie, '')) From ArticuloEspecifico WHERE AEsID = " & rsArt("AEsID")
            Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rsAux.EOF Then oArtIngreso.NroSerie = rsAux(0)
            rsAux.Close
            
        End If
        
        
    ElseIf Not IsNumeric(filter) Then
        
        Cons = "SELECT ArtID, ArtCodigo Codigo, ArtNombre From Articulo WHERE ArtNombre LIKE '" & filter & "%' AND ArtEnUso = 1"
        Set rsArt = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsArt.EOF Then
            oArtIngreso.Articulo = rsArt(0)
            oArtIngreso.ArticuloNombre = rsArt(2)
            rsArt.MoveNext
            If Not rsArt.EOF Then
                
                oArtIngreso.Articulo = 0
                oArtIngreso.ArticuloNombre = ""
                
                Dim oHelp As New clsListadeAyuda
                If oHelp.ActivarAyuda(cBase, Cons, 4000, 1, "B�squeda de art�culos") > 0 Then
                    oArtIngreso.Articulo = oHelp.RetornoDatoSeleccionado(0)
                    oArtIngreso.ArticuloNombre = oHelp.RetornoDatoSeleccionado(2)
                End If
                Set oHelp = Nothing
                
                If oArtIngreso.Articulo = 0 Then
                    rsArt.Close
                    Exit Function
                End If
                
            End If
        End If
        
    End If
    rsArt.Close
    
    If showmsg And oArtIngreso.Articulo = 0 Then
        MsgBox "No existe un art�culo con el filtro ingresado.", vbInformation, "Atenci�n"
    End If
    
    BuscarArticulo = (oArtIngreso.Articulo > 0)
    Exit Function
    
errBA:
    clsGeneral.OcurrioError "Error al buscar el art�culo.", Err.Description, "Buscar art�culos"
    Screen.MousePointer = 0
End Function

Private Sub PresentarCampoIngreso(ByVal campo As String, ByVal foco As Boolean)

    txtArticulo.Visible = (campo = "articulo")
    txtNroSerie.Visible = (campo = "nroserie")
    cboEstado.Visible = (campo = "estado")

    Select Case campo
        
        Case "articulo"
            lblQueEs.Caption = " &Art�culo que ingresa:"
            lblQueEs.BackColor = Me.BackColor
            lblQueEs.ForeColor = vbBlack
            If foco Then txtArticulo.SetFocus
        
        Case "nroserie"
            lblQueEs.Caption = " &N�mero de serie:"
            lblQueEs.BackColor = IIf(oArtIngreso.NroSerie = "", &H8000&, &H80&)
            lblQueEs.ForeColor = vbWhite
            If foco Then txtNroSerie.SetFocus
        
        Case "estado"
            lblQueEs.Caption = " &Estado del art�culo:"
            lblQueEs.BackColor = Me.BackColor
            lblQueEs.ForeColor = vbBlack
            If foco Then cboEstado.SetFocus
            
    End Select
    

End Sub


Private Sub AgregarArticuloEnGrilla()

'Agrego el art�culo a la grilla.
    With lstArticulos
        .AddItem oArtIngreso.ArticuloNombre
        .Cell(flexcpText, .Rows - 1, 1) = oArtIngreso.NroSerie
        .Cell(flexcpText, .Rows - 1, 2) = oArtIngreso.Servicio.LocalRepara.Nombre
        If oArtIngreso.Servicio.Aclaracion <> "" Then
            .Cell(flexcpText, .Rows - 1, 3) = IIf(oArtIngreso.Servicio.Vias > 0, oArtIngreso.Servicio.Vias, "")
        End If
        .Cell(flexcpText, .Rows - 1, 4) = cboEstado.Text
        .Cell(flexcpText, .Rows - 1, 5) = oEstadoIngMerc.ObtenerCadenaEstadosSeleccionados(oArtIngreso.Estados)
        
        'Campos para identificar por si elimina presionando supr el nro de serie pueden ingresar 0 cuando no tienen .
        .Cell(flexcpData, .Rows - 1, 0) = oArtIngreso.Articulo
        .Cell(flexcpData, .Rows - 1, 1) = oArtIngreso.Estados
        .Cell(flexcpData, .Rows - 1, 2) = oArtIngreso.NroSerieArtCambio
        .Cell(flexcpData, .Rows - 1, 3) = oArtIngreso.Servicio.Aclaracion
        .Cell(flexcpData, .Rows - 1, 4) = cboEstado.ListIndex
        
        oArtIngreso.EsRoto = (cboEstado.ItemData(cboEstado.ListIndex) = -1)
        
    End With
    
    'Agrego el art�culo a la colecci�n.
    colRenglonesIngreso.Add oArtIngreso
    
    LimpiarControlesRenglon
    txtArticulo.SetFocus
    
End Sub

Private Function BuscarArticulosParaDevolverDelDocumento() As Boolean
On Error GoTo errBADD
Dim rsArt As rdoResultset
    Dim oRenglon As clsArticuloRenglones
    Cons = "EXEC prg_IngresoMercaderiaCliente_ArticulosPosibles " & Val(hliDocumento.Tag)
    Set rsArt = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsArt.EOF
        
        Set oRenglon = New clsArticuloRenglones
        oRenglon.Articulo = rsArt(0)
        oRenglon.Cantidad = rsArt(1)
        
        colArtsDoc.Add oRenglon
        
        rsArt.MoveNext
    Loop
    rsArt.Close
    BuscarArticulosParaDevolverDelDocumento = True
    Exit Function
    
errBADD:
    clsGeneral.OcurrioError "Error al buscar los art�culos del documento.", Err.Description, "Art�culos del documento"
    
End Function

Public Sub MostrarAyuda(Optional msg As String = "")
    If msg = "" Then
        Select Case Me.ActiveControl.Name
            Case txtDocumento.Name
                lblAyuda.Caption = "B�squeda posible por C.I./R.U.C. del cliente o por c�digo de barras/serie-n�mero del Documento (F12 Vis. Ope.)"
            Case txtArticulo.Name
                lblAyuda.Caption = "Escanee el art�culo o realice la b�squeda por c�digo o nombre"
            Case cboEstado.Name
                lblAyuda.Caption = "Seleccione el estado que indiqu� el estado en que ingresa el/los art�culos"
            Case txtMemo.Name
                lblAyuda.Caption = "Ingrese un comentario."
            Case txtUser.Name
                lblAyuda.Caption = "Ingrese su d�gito de usuario y presione Enter para poder grabar"
            
            Case txtNroSerie.Name
                If oArtIngreso.NroSerie = "" Then
                    lblAyuda.Caption = "Ingrese el n�mero de serie del art�culo que ingrea al local"
                Else
                    lblAyuda.Caption = "Ingrese el n�mero de serie del art�culo que le entrega al cliente"
                End If
                
        End Select
    Else
        lblAyuda.Caption = msg
    End If
End Sub

Private Sub EstadoControlesIngreso(ByVal habilitados As Boolean)
Dim lcolor As Long

    lcolor = IIf(habilitados, vbWindowBackground, vbButtonFace)
    
    With txtArticulo
        .Enabled = habilitados
        .BackColor = lcolor
    End With
        
    With cboEstado
        .Enabled = habilitados
        .BackColor = lcolor
        .Visible = False
    End With
    
    With txtNroSerie
        .Enabled = habilitados
        .BackColor = lcolor
        .Visible = False
    End With
    
    With txtMemo
        .Enabled = habilitados
        .BackColor = lcolor
    End With
    
    With txtUser
        .Enabled = habilitados
        .BackColor = lcolor
    End With
    
    lstArticulos.Enabled = habilitados
    
    cboAccion.Enabled = Not habilitados
    
    butAceptar.Enabled = habilitados
    
End Sub

Private Sub LimpiarControlesRenglon()
    txtArticulo.Text = "": txtArticulo.Tag = ""
    txtNroSerie.Text = ""
    cboEstado.Text = ""
    Set oArtIngreso = Nothing
    PresentarCampoIngreso "articulo", False
End Sub

Private Sub LimpiarControlesArticulos()
    
    LimpiarControlesRenglon
    txtMemo.Text = ""
    txtUser.Text = "": txtUser.Tag = 0
    lstArticulos.Rows = 1
    Set colRenglonesIngreso = Nothing
    Set colRenglonesIngreso = New Collection
    
End Sub

Private Sub LimpiarControlesDocumento()
    hliCliente.Caption = ""
    hliDocumento.Caption = ""
    hliCliente.Tag = ""
    hliDocumento.Tag = ""
    Set colArtsDoc = Nothing
    Set colArtsDoc = New Collection
End Sub

Private Sub butAceptar_Click()
    
    On Error Resume Next
    If colRenglonesIngreso.Count = 0 Then
        MsgBox "No hay datos ingresados.", vbExclamation, "Validaci�n"
        txtArticulo.SetFocus
        Exit Sub
    End If
    
    If Val(txtUser.Tag) = 0 Then
        MsgBox "Ingrese su d�gito de usuario.", vbExclamation, "Validaci�n"
        txtUser.SetFocus
        Exit Sub
    End If
    
    'Si es un remito recepci�n tengo que obligar a ingresar todos los art�culos.
    If MsgBox("�Confirma almacenar los datos ingresados?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
        GrabarIngreso
    End If
    
End Sub

Private Sub butCancelar_Click()
    
    'cancelo el ingreso.
    LimpiarControlesDocumento
    LimpiarControlesArticulos
    EstadoControlesIngreso False
    txtDocumento.Text = ""
    txtDocumento.SetFocus
    
End Sub

Private Sub cboAccion_Click()
    lblQueEs.Caption = " &Art�culo que ingresa"
End Sub

Private Sub cboAccion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtDocumento.SetFocus
End Sub

Private Sub cboEstado_GotFocus()
    With cboEstado
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    MostrarAyuda
    If oArtIngreso Is Nothing Then Exit Sub
    If ((InStr(1, "," & paArticuloARoto & ",", "," & oArtIngreso.Articulo & ",") > 0) And Val(cboEstado.Tag) <> 1) Or _
        ((InStr(1, "," & paArticuloARoto & ",", "," & oArtIngreso.Articulo & ",") = 0) And Val(cboEstado.Tag) <> 0) Then
    
        CargoComboEstado (InStr(1, "," & paArticuloARoto & ",", "," & oArtIngreso.Articulo & ",") > 0)
    End If
End Sub

Private Sub cboEstado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        LimpiarControlesRenglon
        txtArticulo.SetFocus
    End If
End Sub

Private Sub cboEstado_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        
        'Validaci�n de datos.
        If Not (oArtIngreso.Articulo > 0 And cboEstado.ListIndex > -1) Then
            MsgBox "Faltan ingresar datos para cargar el art�culo.", vbExclamation, "Validaci�n"
            Exit Sub
        End If
        
        If oEstadoIngMerc.Estados.Count = 0 Then oEstadoIngMerc.CargarEstados cBase
        
        'Si el estado tiene en el item data = 0 --> tengo que pedir los estados.
        If Val(cboEstado.ItemData(cboEstado.ListIndex)) = 0 Then
            
            If colLocalesRepara Is Nothing Then
                Set colLocalesRepara = New Collection
                CargarLocalesRepara
            End If
            
            butAceptar.Enabled = False
            butCancelar.Enabled = False
            
            Dim frmEstSer As New frmEstadoServicio
            With frmEstSer
            
                Set .oEstadoIngMerc = oEstadoIngMerc
                Set .Servicio = oArtIngreso.Servicio
                .EstadosSeleccionados = oArtIngreso.Estados
                '.Estado = oArtIngreso.Estado
                
                .IDArticulo = oArtIngreso.Articulo
                
                Dim mRect As RECT
                GetWindowRect cboEstado.hwnd, mRect
                .Left = mRect.Left * 15
                .Top = (mRect.Top * 15) + cboEstado.Height
                
                .Show vbModal
                
                If .EstadosSeleccionados <> "" Then
                    oArtIngreso.Estados = .EstadosSeleccionados
                    Set oArtIngreso.Servicio = .Servicio
                End If
                
            End With
            Unload frmEstSer
            Set frmEstSer = Nothing
            
            butAceptar.Enabled = True
            butCancelar.Enabled = True
            
            'If oArtIngreso.Estado = 0 Then Exit Sub
            If oArtIngreso.Estados = "" Then Exit Sub
            
        Else
            oArtIngreso.Estados = "" 'cboEstado.ItemData(cboEstado.ListIndex)
        End If
        
        If cboAccion.ItemData(cboAccion.ListIndex) = 1 Then
            
            AgregarArticuloEnGrilla
        
        Else
        
            txtNroSerie.Text = ""
            'Pido el nro. de serie del que entrega.
            PresentarCampoIngreso "nroserie", True
            
        End If
        
    End If
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
Dim sPaso As String

    sPaso = "Controles"
    EstadoControlesIngreso False
    LimpiarControlesArticulos
    LimpiarControlesDocumento
    
    'ubico los controles
    cboEstado.Move txtArticulo.Left, txtArticulo.Top
    txtNroSerie.Move txtArticulo.Left, txtArticulo.Top
        
    'Cargo combo t�tulo.
    With cboAccion
        .Clear
        .AddItem "por devoluci�n de mercader�a"
        .ItemData(.NewIndex) = TAE_Devolucion
        .AddItem "por cambio de mercader�a"
        .ItemData(.NewIndex) = TAE_Cambio
        .ListIndex = 0
    End With
    
'    With txtArticulo
'        Set .Connect = cBase
'        .KeyQuerySP = "IngresoPorDevolucion"
'        .DisplayCodigoArticulo = True
'    End With
    
    With lstArticulos
        .Rows = 1: .Cols = 1
        
        .RowHeight(0) = 315
        .RowHeightMin = 285
        
        .FormatString = "<Art�culo|<Serie|<Reparar en|>V�as|Estado|<Detalle estados"
        
        .ColWidth(0) = 2600
        .ColWidth(1) = 1400
        .ColWidth(2) = 1000
        .ColWidth(4) = 1200
        .ColWidth(5) = 1500
        .ExtendLastCol = True
        
    End With
    
    sPaso = "Posici�n y Vista"
    CargoComboEstado False
    
    Dim oAC As New clsConexion
    sConnect = oAC.TextoConexion("Comercio")
    Set oAC = Nothing

    
    Exit Sub
errLoad:
    clsGeneral.OcurrioError "Error al inicializar el formulario", sPaso & vbCrLf & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    
    Set colLocalesRepara = Nothing
    Set colArtsDoc = Nothing
    
    Set oEstadoIngMerc = Nothing
    cBase.Close
    Set cBase = Nothing
    Set clsGeneral = Nothing
    
    Dim objFnc As New clsFncGlobales
    objFnc.SetPositionForm Me
    Set objFnc = Nothing
    
End Sub

Private Sub hliDocumento_Click()
On Error Resume Next
    If Val(hliDocumento.Tag) > 0 Then
        Shell App.Path & "\detalle de factura.exe " & hliDocumento.Tag, vbNormalFocus
    End If
End Sub

Private Sub Label5_DblClick()
'    TestPrint
End Sub

Private Sub lstArticulos_GotFocus()
    lblAyuda.Caption = "Art�culos que ingresan al local. (<Supr> eliminar, <Espacio> edita estado)"
End Sub

Private Sub lstArticulos_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If lstArticulos.Row < 1 Then Exit Sub
    
    Dim iQ As Byte
    Select Case KeyCode
        Case vbKeyDelete
            'Elimino de la grilla y de la colecci�n.
            Dim oArt As clsRenglonIngreso
            With lstArticulos
                
                For iQ = 1 To colRenglonesIngreso.Count
                    
                    Set oArt = colRenglonesIngreso(iQ)
                    If oArt.Articulo = .Cell(flexcpData, .Row, 0) And oArt.NroSerie = .Cell(flexcpText, .Row, 1) And _
                        oArt.Estados = .Cell(flexcpData, .Row, 1) And oArt.NroSerieArtCambio = .Cell(flexcpData, .Row, 2) And _
                        oArt.Servicio.Aclaracion = .Cell(flexcpData, .Row, 3) And oArt.Servicio.Vias = Val(.Cell(flexcpText, .Row, 3)) Then
                        
                        Set oArt = Nothing
                        colRenglonesIngreso.Remove iQ
                        
                        lstArticulos.RemoveItem lstArticulos.Row
                        Exit Sub
                        
                    End If
                    
                Next
            End With
            
        Case vbKeySpace
            If Not oArtIngreso Is Nothing Then
                MsgBox "Para poder editar un registro en la grilla debe cancelar el ingreso actual.", vbExclamation, "Validaci�n"
                Exit Sub
            End If
            
            With lstArticulos
                
                For iQ = 1 To colRenglonesIngreso.Count
                    
                    Set oArt = colRenglonesIngreso(iQ)
                    If oArt.Articulo = .Cell(flexcpData, .Row, 0) And oArt.NroSerie = .Cell(flexcpText, .Row, 1) And _
                        oArt.Estados = .Cell(flexcpData, .Row, 1) And oArt.NroSerieArtCambio = .Cell(flexcpData, .Row, 2) And _
                        oArt.Servicio.Aclaracion = .Cell(flexcpData, .Row, 3) And oArt.Servicio.Vias = Val(.Cell(flexcpText, .Row, 3)) Then
                        
                        Set oArtIngreso = oArt
                        Exit For
                        
                    End If
                    
                Next
            End With
            
            If oArtIngreso Is Nothing Then
                MsgBox "Error en los datos, reintente.", vbExclamation, "Atenci�n"
                Exit Sub
            Else
                cboEstado.Text = lstArticulos.Cell(flexcpText, lstArticulos.Row, 4)
                cboEstado.ListIndex = lstArticulos.Cell(flexcpData, lstArticulos.Row, 4)
                txtArticulo.Tag = oArtIngreso.Articulo
                PresentarCampoIngreso "estado", True
                lstArticulos.RemoveItem lstArticulos.Row
            End If
    End Select
End Sub

Private Sub lstArticulos_LostFocus()
    lblAyuda.Caption = ""
End Sub

Private Sub txtArticulo_Change()
    If Val(txtArticulo.Tag) > 0 Then
        txtArticulo.Tag = "": Set oArtIngreso = Nothing
    End If
End Sub

Private Sub txtArticulo_GotFocus()
    With txtArticulo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    MostrarAyuda
End Sub

Private Sub txtArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        LimpiarControlesRenglon
        txtArticulo.SetFocus
    End If
End Sub

Private Sub txtArticulo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        
        If Trim(txtArticulo.Text) = "" And lstArticulos.Rows > 1 Then
            txtMemo.SetFocus
        Else
            
            If Val(txtArticulo.Tag) = 0 And Trim(txtArticulo.Text) <> "" Then
                
                'Cargo el art�culo.
                
                If Not BuscarArticulo(True, txtArticulo.Text) Then
                    txtArticulo_GotFocus
                    Exit Sub
                End If
                
                If Val(hliDocumento.Tag) = 0 Then
                    'Busco el documento para el cliente y para el art�culo
                    CargoPosibleFactura oArtIngreso.Articulo
                    
                    If Val(hliDocumento.Tag) > 0 Then
                        BuscarArticulosParaDevolverDelDocumento
                        If colArtsDoc.Count = 0 Then
                            MsgBox "No hay art�culos posibles para devolver en ese documento.", vbExclamation, "ATENCI�N"
                            Screen.MousePointer = 0
                            Exit Sub
                        End If
                    End If
                End If
                
                'Si no tengo documento no dejo seguir.
                If (Val(hliDocumento.Tag) = 0) Then
                    MsgBox "No hay una factura para el art�culo ingresado, no puede continuar.", vbExclamation, "Posible error"
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                
                If colArtsDoc.Count > 0 Then
                    
                    Dim iCant As Integer: iCant = 0
                    'Verifico que tenga el art�culo en la colecci�n.
                    Dim oArticulo As clsArticuloRenglones
                    For Each oArticulo In colArtsDoc
                        If oArticulo.Articulo = oArtIngreso.Articulo Then
                            iCant = oArticulo.Cantidad
                            Exit For
                        End If
                    Next
                    
                    If iCant = 0 Then
                        MsgBox "No es posible cargar el art�culo ingresado.", vbExclamation, "Validaci�n"
                        txtArticulo_GotFocus
                        Exit Sub
                    Else
                        
                        'Ahora sumo los que ingres�.
                        Dim oArts As clsRenglonIngreso
                        For Each oArts In colRenglonesIngreso
                            If oArts.Articulo = oArtIngreso.Articulo Then
                                iCant = iCant - 1
                            End If
                        Next
                        
                        If iCant = 0 Then
                            MsgBox "Ya complet� la cantidad posible para este art�culo, verifique.", vbExclamation, "Posible error"
                            txtArticulo_GotFocus
                            Exit Sub
                        End If
                    End If
                    
                    
                End If
                
                'Si es nro. de serie verifico que no est� ingresado.
                If oArtIngreso.NroSerie <> "" Then
                    
                    Dim oArt As clsRenglonIngreso
                    For Each oArt In colRenglonesIngreso
                        If oArt.NroSerie = oArtIngreso.NroSerie Or oArt.NroSerie = oArtIngreso.NroSerieArtCambio Then
                            MsgBox "Ya existe un art�culo ingresado con ese n�mero de serie.", vbExclamation, "Posible Duplicaci�n"
                            txtArticulo_GotFocus
                            Exit Sub
                        End If
                    Next
                End If
                
                txtArticulo.Text = oArtIngreso.ArticuloNombre
                txtArticulo.Tag = oArtIngreso.Articulo
                
                'Presento campo siguiente
                If oArtIngreso.NroSerie <> "" Then
                    PresentarCampoIngreso "estado", True
                Else
                    PresentarCampoIngreso "nroserie", True
                End If
                
            End If
        End If
    End If
End Sub

Private Sub txtArticulo_LostFocus()
    lblAyuda.Caption = ""
End Sub

Private Sub txtDocumento_Change()
    If Val(hliDocumento.Tag) > 0 Or Val(hliCliente.Tag) > 0 Then
        LimpiarControlesDocumento
        LimpiarControlesArticulos
        EstadoControlesIngreso False
    End If
End Sub

Private Sub txtDocumento_GotFocus()
    With txtDocumento
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    MostrarAyuda
End Sub

Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF12
            Shell App.Path & "\voperaciones.exe " & Val(hliCliente.Tag)
    End Select
End Sub

Private Sub txtDocumento_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn And Trim(txtDocumento.Text) <> "" Then
        
        If Val(hliDocumento.Tag) > 0 Or Val(hliCliente.Tag) > 0 Then
            If txtArticulo.Enabled And txtArticulo.Visible Then txtArticulo.SetFocus
            Exit Sub
        End If
        
        Dim objHelp As clsListadeAyuda
        On Error GoTo errBD
        Screen.MousePointer = 11
        
        
        Cons = "EXEC prg_BuscarCliente 1, '', '', '', '', '', '" & txtDocumento.Text & "', 0, 0, '', '', 16"
        Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If Not rsAux.EOF Then
            
            If Not IsNull(rsAux("CliCodigo")) Then
                
                If rsAux("CliCodigo") > 0 Then
                    
                    If rsAux("DocCodigo") > 0 Then
                        hliDocumento.Tag = rsAux("DocCodigo")
                        hliDocumento.Caption = rsAux("Documento")
                        lblDocumento.Tag = rsAux("DocFModificacion")
                    End If
                    
                    hliCliente.Caption = "(" & RTrim(rsAux("C.I./R.U.C.")) & ") " & rsAux("Cliente")
                    hliCliente.Tag = rsAux("CliCodigo")
                    rsAux.MoveNext
                    If Not rsAux.EOF Then
                        rsAux.Close
                        
                        hliDocumento.Tag = ""
                        hliDocumento.Caption = ""
                        hliCliente.Caption = ""
                        hliCliente.Tag = ""
                        lblDocumento.Tag = ""
                                                
                        'Abro lista de ayuda.
                        Set objHelp = New clsListadeAyuda
                        If objHelp.ActivarAyuda(cBase, Cons, 5000, 3, "B�squeda") > 0 Then
                            
                            hliDocumento.Tag = objHelp.RetornoDatoSeleccionado(0)
                            hliCliente.Tag = objHelp.RetornoDatoSeleccionado(1)
                                                        
                        End If
                    Else
                        rsAux.Close
                    End If
                    
                End If
            Else
                rsAux.Close
            End If
            
        Else
        
            MsgBox "No hay resultados para el dato ingresado.", vbInformation, "B�squeda"
            rsAux.Close
            
        End If
        
        If Val(hliDocumento.Tag) > 0 Then
            
            Dim TipoDoc As Byte
            Cons = "SELECT DocTipo FROM Documento WHERE DocCodigo = " & Val(hliDocumento.Tag)
            Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            TipoDoc = rsAux("DocTipo")
            rsAux.Close
            
            If TipoDoc = 6 Then
                 Cons = "SELECT RDoDocumento FROM RemitoDocumento WHERE RDoRemito = " & Val(hliDocumento.Tag)
                Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not rsAux.EOF Then hliDocumento.Tag = rsAux(0) Else hliDocumento.Tag = 0
                rsAux.Close
                
                If Val(hliDocumento.Tag) = 0 Then
                    MsgBox "Ud. escaneo un remito que no tiene asociado ning�n documento de compra.", vbExclamation, "ATENCI�N"
                    Exit Sub
                End If
                
            End If
            
            
            
            'Cargo a partir de un documento
            If hliDocumento.Caption = "" Then
                
                'es seleccionado por lista de ayuda.
                Cons = "SELECT DocCodigo, CliCodigo, DocFModificacion, dbo.NombreTipoDocumento(100+DocTipo) + ' ' + rTrim(DocSerie)+Convert(VarChar(10), DocNumero) Documento," & _
                    " RTrim(IsNull(CEmFantasia, rTrim(CPeApellido1) + ', ' + RTrim(CPeNombre1))) Cliente, ISNULL(CliCiRUC, '') [C.I./R.U.C.]" & _
                    " FROM Documento INNER JOIN CLiente ON DocCliente = CliCodigo" & _
                    " LEFT OUTER JOIN CPersona ON CPeCliente = CliCodigo" & _
                    " LEFT OUTER JOIN CEmpresa ON CEmCliente = CliCodigo" & _
                    " WHERE DocCodigo = " & hliDocumento.Tag
                    
                Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not rsAux.EOF Then
                    hliDocumento.Tag = rsAux("DocCodigo")
                    hliDocumento.Caption = rsAux("Documento")
                    hliCliente.Caption = "(" & RTrim(rsAux("C.I./R.U.C.")) & ") " & rsAux("Cliente")
                    hliCliente.Tag = rsAux("CliCodigo")
                    lblDocumento.Tag = rsAux("DocFModificacion")
                End If
                rsAux.Close
                
            End If
            
        ElseIf Val(hliCliente.Tag) > 0 Then
            
            'Cargo a partir de un cliente.
            'Si estoy buscando para cambio de productos entonces le pido que seleccione un documento.
            If cboAccion.ItemData(cboAccion.ListIndex) = 2 Then
                
                'Busco los documentos que tengan art�culos entregados.
                Cons = "SELECT DocCodigo, DocFModificacion, dbo.NombreTipoDocumento(100+DocTipo) + ' ' + rTrim(DocSerie)+Convert(VarChar(10), DocNumero) Documento, dbo.ListaArticulosDelDocumento(DocCodigo) Art�culos" & _
                    " FROM ((Documento INNER JOIN Renglon ON RenDocumento = DocCodigo)" & _
                    " INNER JOIN Articulo ON ArtID = RenArticulo AND ArtTipo <> 151)" & _
                    " WHERE DocTipo IN (1,2,6) AND RenCantidad <> RenARetirar AND DocCliente = " & hliCliente.Tag
                    
                Set objHelp = New clsListadeAyuda
                If objHelp.ActivarAyuda(cBase, Cons, 5000, 2, "B�squeda") > 0 Then
                    hliDocumento.Tag = objHelp.RetornoDatoSeleccionado(0)
                    hliDocumento.Caption = objHelp.RetornoDatoSeleccionado(2)
                    lblDocumento.Tag = objHelp.RetornoDatoSeleccionado(1)
                Else
                    'no permito seguir con el ingreso.
                    hliCliente.Tag = ""
                    Set objHelp = Nothing
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                Set objHelp = Nothing
                
                If hliDocumento.Tag = "" Then
                    MsgBox "No hay un documento para poder realizar cambio de productos.", vbInformation, "Cambio de producto"
                End If
            Else
                'No retorna por documento es crear una ficha de devoluci�n (remito recepci�n).
            End If
        End If
        
        If Val(hliCliente.Tag) > 0 Then
            BuscoComentariosAlerta Val(hliCliente.Tag), True
        End If
        
    
        'Si tengo cliente o documento asignado
        If Val(hliDocumento.Tag) > 0 Or Val(hliCliente.Tag) > 0 Then
            
            If Val(hliDocumento.Tag) > 0 Then
                
                'Busco si el documento posee art�culos disponibles para devolver.
                BuscarArticulosParaDevolverDelDocumento
                If colArtsDoc.Count = 0 Then
                    MsgBox "Atenci�n el documento no posee art�culos entregados o del mismo no se pueden devolver m�s art�culos.", vbInformation, "ATENCI�N"
                Else
                    EstadoControlesIngreso True
                    txtArticulo.SetFocus
                End If
            
            ElseIf cboAccion.ItemData(cboAccion.ListIndex) = 1 Then
                    
                    EstadoControlesIngreso True
                    txtArticulo.SetFocus
                    
            End If
            
            If txtArticulo.Enabled Then
                BuscarCuotasVencidasCliente hliCliente.Tag, hliCliente.Caption, True
            End If
            
        End If
        
    End If
    Screen.MousePointer = 0
    Exit Sub
errBD:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al buscar.", Err.Description, "B�squeda"
    
End Sub

Private Sub txtDocumento_LostFocus()
    lblAyuda.Caption = ""
End Sub

Private Sub txtMemo_GotFocus()
    MostrarAyuda
    With txtMemo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtMemo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtUser.SetFocus
End Sub

Private Sub txtMemo_LostFocus()
lblAyuda.Caption = ""
End Sub

Private Sub txtNroSerie_GotFocus()
    MostrarAyuda
    With txtNroSerie
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtNroSerie_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        LimpiarControlesRenglon
        txtArticulo.SetFocus
    End If
End Sub

Private Sub txtNroSerie_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And Trim(txtNroSerie.Text) <> "" Then
        
        'verifico que no este ingresado el art�culo
        If ExisteNroSerieEnGrilla(txtNroSerie.Text) Then
            MsgBox "Ya ingres� un art�culo con ese n�mero de serie, verifiqu�.", vbExclamation, "Validaci�n"
            Exit Sub
        End If
        
        If oArtIngreso.NroSerie = "" Then
            
            oArtIngreso.NroSerie = txtNroSerie.Text
            PresentarCampoIngreso "estado", True
        
        ElseIf cboAccion.ItemData(cboAccion.ListIndex) = 2 Then
            
            If oArtIngreso.NroSerie = txtNroSerie.Text And txtNroSerie.Text <> "0" Then
                MsgBox "El n�mero de serie ingresado es el mismo que del art�culo que ingresa.", vbExclamation, "Validaci�n"
                Exit Sub
            End If
            
            oArtIngreso.NroSerieArtCambio = txtNroSerie.Text
            'Agrego el art�culo a la lista.
            AgregarArticuloEnGrilla
            
        End If
        
    End If
    
End Sub

Private Sub txtUser_Change()
    txtUser.Tag = ""
End Sub

Private Sub txtUser_GotFocus()
    MostrarAyuda
    With txtUser
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And IsNumeric(txtUser.Text) Then
        If Val(txtUser.Tag) = 0 Then
            Dim objFnc As New clsFncGlobales
            txtUser.Tag = objFnc.BuscarUsuario(CInt(txtUser.Text))
            Set objFnc = Nothing
        End If
        
        If Val(txtUser.Tag) > 0 Then
            butAceptar_Click
        End If
        
    End If
End Sub

Private Sub txtUser_LostFocus()
    lblAyuda.Caption = ""
End Sub

Private Function CargoPosibleFactura(ByVal IDArticulo As Long) As Long

    Cons = "SELECT DocCodigo, DocFModificacion, dbo.NombreTipoDocumento(100+DocTipo) + ' ' + rTrim(DocSerie)+Convert(VarChar(10), DocNumero) Documento, DocFecha Fecha, dbo.ListaArticulosDelDocumento(DocCodigo) Art�culos" & _
        " FROM ((Documento INNER JOIN Renglon ON RenDocumento = DocCodigo And (RenArticulo = " & IDArticulo & " OR " & IDArticulo & " = 0))" & _
        " INNER JOIN Articulo ON ArtID = RenArticulo AND ArtTipo <> 151)" & _
        " WHERE DocTipo IN (1,2,6) AND RenCantidad <> RenARetirar AND DocCliente = " & hliCliente.Tag
        
    Dim objHelp As New clsListadeAyuda
    objHelp.CerrarSiEsUnico = True
    If objHelp.ActivarAyuda(cBase, Cons, 5000, 2, "B�squeda") > 0 Then
        hliDocumento.Tag = objHelp.RetornoDatoSeleccionado(0)
        hliDocumento.Caption = objHelp.RetornoDatoSeleccionado(2)
        lblDocumento.Tag = objHelp.RetornoDatoSeleccionado(1)
        If Format(objHelp.RetornoDatoSeleccionado(3), "dd/MM/yyyy") = Date Then
            
        End If
    End If
    Set objHelp = Nothing
    CargoPosibleFactura = Val(hliDocumento.Tag)
    
End Function

Public Sub BuscoComentariosAlerta(idCliente As Long, _
                                                   Optional Alerta As Boolean = False, Optional Cuota As Boolean = False, _
                                                   Optional Decision As Boolean = False, Optional Informacion As Boolean = False)
                                                   
Dim rsCom As rdoResultset
Dim aCom As String
Dim sHay As Boolean

    On Error GoTo errMenu
    Screen.MousePointer = 11
    sHay = False
    'Armo el str con los comentarios a consultar-------------------------------------------------
    If Not Alerta And Not Cuota And Not Decision And Not Informacion Then Exit Sub
    aCom = ""
    If Alerta Then aCom = aCom & Accion.Alerta & ", "
    If Cuota Then aCom = aCom & Accion.Cuota & ", "
    If Decision Then aCom = aCom & Accion.Decision & ", "
    If Informacion Then aCom = aCom & Accion.Informacion & ", "
    aCom = Mid(aCom, 1, Len(aCom) - 2)
    '---------------------------------------------------------------------------------------------------
    
    Cons = "Select * From Comentario, TipoComentario " _
            & " Where ComCliente = " & idCliente _
            & " And ComTipo = TCoCodigo " _
            & " And TCoAccion IN (" & aCom & ")"
    Set rsCom = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not rsCom.EOF Then sHay = True
    rsCom.Close
    
    If sHay Then
        Dim aObj As New clsCliente
        aObj.Comentarios idCliente:=idCliente
        DoEvents
        Set aObj = Nothing
    End If
    MsgClienteNoVender idCliente, True
    Screen.MousePointer = 0
    Exit Sub
    
errMenu:
    clsGeneral.OcurrioError "Ocurri� un error al acceder al fomulario de comentarios.", Err.Description
    Screen.MousePointer = 0
End Sub

Public Function MsgClienteNoVender(ByVal iCliente As Long, ByVal bShowMsg As Boolean) As Boolean
Dim rsCom As rdoResultset
    MsgClienteNoVender = False
    Set rsCom = cBase.OpenResultset("exec gennovender " & iCliente, rdOpenDynamic, rdConcurValues)
    If Not rsCom.EOF Then
        If Not IsNull(rsCom(0)) Then
            If rsCom(0) = 1 Then
                MsgClienteNoVender = True
                If bShowMsg Then
                    Screen.MousePointer = 0
                    MsgBox "Atenci�n: el cliente tiene la categor�a de no vender. Consultar con gerencia!", vbCritical, "ATENCI�N"
                End If
            End If
        End If
    End If
    rsCom.Close
End Function

Private Sub CargoComboEstado(ByVal admiteRoto As Boolean)
    'Cargo los estados en el combo y ubico el formulario.
    Dim objFnc As New clsFncGlobales
    objFnc.GetPositionForm Me
    Me.Width = 8535
    Me.Height = 6510
    
    Dim sQy As String
    sQy = "SELECT IsNull(CodValor1, 0), rTrim(CodTexto)" & _
                    " FROM Codigos WHERE codCual = 144 "
    If admiteRoto Then
        sQy = "SELECT -1, 'Roto' UNION ALL " & sQy
    Else
        sQy = sQy & "Order By CodID"
    End If
    objFnc.CargoCombo sQy, cboEstado
    Set objFnc = Nothing
    
    cboEstado.Tag = IIf(admiteRoto, "1", "0")
    
End Sub
