VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmAnularVentas 
   Caption         =   "Eliminar Ventas telefónicas y redpagos"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9315
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAnularVentas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMenu 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   9285
      TabIndex        =   1
      Top             =   0
      Width           =   9315
      Begin VB.TextBox txtDias 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "8"
         ToolTipText     =   "Ventas menores a los días indicados"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton butConsultar 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   5400
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton butGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   6600
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Días:"
         Height          =   255
         Left            =   4440
         TabIndex        =   5
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione que ventas desea anular"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3855
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7455
      _cx             =   13150
      _cy             =   4683
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
      FixedCols       =   1
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
      AutoSearchDelay =   2
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
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmAnularVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AnuloVenta(ByVal IDVenta As Long, ByVal Usuario As Long)
On Error GoTo ErrAE

Screen.MousePointer = 11

    cBase.BeginTrans
    On Error GoTo ErrResumo
        
    Dim rsVta As rdoResultset
    Cons = "Select * From VentaTelefonica Where VTeCodigo = " & IDVenta & _
        " AND VTeAnulado IS NULL and VTeTipo IN (7, 32, 33, 44) AND VTeDocumento IS NULL "
    Set rsVta = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    'Se la pelaron.
    If rsVta.EOF Then
        cBase.RollbackTrans
        rsVta.Close
        Screen.MousePointer = vbDefault
        MsgBox "Otra terminal elimino la venta, verifique.", vbExclamation, "ATENCIÓN"
    End If
        
    'Borro si tiene envíos.-------------------------------
    'Para cada artículo que este en envío le hago el movimiento de stock
    Cons = "Select * From RenglonEnvio Where REvEnvio IN (" _
        & "Select EnvCodigo From Envio Where EnvTipo = " & TipoEnvio.Cobranza _
        & " And EnvDocumento = " & rsVta!VTeCodigo & ")" _
        & " And REvArticulo Not IN (Select ArtID From Articulo Where ArtTipo = " & paTipoArticuloServicio & ")"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
    
        'IMPORTANTE YA RESTO EL ARETIRAR ya que no updateo el renglón de la venta telefónica.
        MarcoStockVenta Usuario, RsAux!REvArticulo, 0, RsAux!REvAEntregar * -1, 0, TipoDocumento.ContadoDomicilio, rsVta!VTeCodigo, paCodigoDeSucursal
        RsAux.MoveNext
    Loop
    RsAux.Close
        
    Cons = " Delete RenglonEnvio Where REvEnvio IN (" _
        & "Select EnvCodigo From Envio Where EnvTipo = " & TipoEnvio.Cobranza _
        & " And EnvDocumento = " & rsVta!VTeCodigo & ")"
    cBase.Execute (Cons)
    
    Cons = " Delete Envio Where  EnvTipo = " & TipoEnvio.Cobranza _
        & " And EnvDocumento = " & rsVta!VTeCodigo
    cBase.Execute (Cons)
    
    Cons = "Delete Direccion Where DirCodigo IN(" _
        & "Select EnvDireccion From Envio Where EnvTipo = " & TipoEnvio.Cobranza _
        & " And EnvDocumento = " & rsVta!VTeCodigo & ")"
    cBase.Execute (Cons)
        
    FechaDelServidor
    'Marco el stock.
    Cons = "Select * From RenglonVtaTelefonica " _
        & " Where RVTVentaTelefonica = " & rsVta!VTeCodigo _
        & " And RVTArticulo Not IN (Select ArtID From Articulo Where ArtTipo = " & paTipoArticuloServicio & ")"
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    Do While Not RsAux.EOF
        If RsAux!RVTARetirar > 0 Then
            MarcoStockVenta Usuario, RsAux!RVtArticulo, RsAux!RVTARetirar * -1, 0, 0, TipoDocumento.ContadoDomicilio, rsVta!VTeCodigo, paCodigoDeSucursal
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
        
    'Instalaciones la anulo.
    Cons = "Select * from Instalacion Where InsTipoDocumento = 2 And InsDocumento = " & rsVta!VTeCodigo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        RsAux.Edit
        RsAux!InsAnulada = Format(gFechaServidor, "yyyy-mm-dd hh:nn")
        RsAux!InsFechaModificacion = Format(gFechaServidor, "yyyy/mm/dd hh:mm:ss")
        RsAux.Update
    End If
    RsAux.Close
    
    'where instipodocumento = 2
    '................................................
    
    rsVta.Edit
    rsVta!VTeAnulado = Format(gFechaServidor, sqlFormatoFH)
    rsVta.Update
    rsVta.Close
    
    cBase.CommitTrans
    
    Screen.MousePointer = vbDefault
    Exit Sub

ErrAE:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrio un error al intentar eliminar la venta."
    Exit Sub
    
ErrResumo:
    Resume ErrTrans
    
ErrTrans:
    cBase.RollbackTrans
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Error al intentar eliminar la venta.", Err.Description
    rsVta.Requery

End Sub

Private Sub Consultar()
On Error GoTo errQuery
    
    If Not IsNumeric(txtDias.Text) Then
        MsgBox "Ingrese la cantidad de días a descartar en la consulta menores a hoy.", vbExclamation, "ATENCIÓN"
        txtDias.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    vsGrid.Rows = 1
    
    Cons = "SELECT VTeCodigo, VTeFechaLlamado Fecha, Case WHEN VTeTipo = 44 THEN 'Redpagos' WHEN VTeTipo = 32 THEN 'WEB AC' WHEN VTeTipo =  33 THEN 'Web' WHEN VTeTipo = 7 THEN 'Telefónica' END Tipo " & _
            " From VentaTelefonica" & _
            " WHERE VTeAnulado IS NULL and VTeTipo IN (7, 32, 33, 44) AND VTeDocumento IS NULL and VTeFechaLlamado < DATEADD(d, -" & txtDias.Text & ", GetDATE()) " & _
            " ORDER BY VteCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        With vsGrid
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RsAux("VTeCodigo")
            .Cell(flexcpText, .Rows - 1, 2) = RsAux("Fecha")
            .Cell(flexcpText, .Rows - 1, 3) = RsAux("Tipo")
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub
errQuery:
    clsGeneral.OcurrioError "Error al consultar", Err.Description, "Anular Ventas"
End Sub
Private Sub butConsultar_Click()
    Consultar
End Sub

Private Sub butGrabar_Click()
    If MsgBox("¿Desea anular la(s) venta(s) seleccionada(s)?", vbQuestion + vbYesNo, "ELIMINAR") = vbYes Then
    
        Dim strFecha As String, strUsuario As String
        strUsuario = vbNullString
        strUsuario = InputBox("Ingrese su código de usuario.", "Emitir Factura")
        If Trim(strUsuario) = vbNullString Then
            MsgBox "No se almacenará la información.", vbInformation, "ATENCIÓN"
            Exit Sub
        Else
            If Not IsNumeric(strUsuario) Then
                MsgBox "El formato ingresado no es numérico.", vbExclamation, "ATENCIÓN"
                Exit Sub
            Else
                strUsuario = GetUIDCodigo(CLng(strUsuario))
            End If
        End If
    
        Dim iRow As Integer
        For iRow = 1 To vsGrid.Rows - 1
            If vsGrid.Cell(flexcpChecked, iRow, 0) = flexChecked Then
                AnuloVenta vsGrid.Cell(flexcpText, iRow, 1), CLng(strUsuario)
            End If
        Next
    End If
    
    Consultar
End Sub

Private Sub Form_Load()
On Error Resume Next
    With vsGrid
        .Rows = 1
        .Cols = 1
        .Editable = True
        .FormatString = "Anular|Venta|Fecha|Tipo|"
        .FixedCols = 0
        .ColDataType(0) = flexDTBoolean
        .ColWidth(1) = 1200
        .ColWidth(2) = 1500
        .ColWidth(3) = 2000
        .ExtendLastCol = True
    End With
End Sub

Private Sub Form_Resize()
On Error Resume Next
    vsGrid.Move 0, picMenu.Height, ScaleWidth, ScaleHeight - picMenu.Height
End Sub

Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 Cancel = (Col <> 0)
End Sub

Private Function GetUIDCodigo(ByVal Digito As Long) As Long
Dim RsUsr As rdoResultset
On Error GoTo ErrBUD
    GetUIDCodigo = 0
    Cons = "Select * from Usuario Where UsuDigito = " & Digito
    Set RsUsr = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsUsr.EOF Then GetUIDCodigo = RsUsr!UsuCodigo
    RsUsr.Close
    Exit Function
ErrBUD:
    MsgBox "Error al buscar el usuario." & vbCr & vbCr & "Error: " & Err.Description, vbCritical, "ATENCIÓN"
End Function

