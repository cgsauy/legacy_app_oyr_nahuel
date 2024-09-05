VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form frmEventos 
   Caption         =   "Eventos"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEventos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VSFlex6DAOCtl.vsFlexGrid vsGrid 
      Height          =   2235
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3942
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
      BackColor       =   15792885
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483636
      ForeColorFixed  =   -2147483634
      BackColorSel    =   128
      ForeColorSel    =   16777215
      BackColorBkg    =   15790320
      BackColorAlternate=   15793661
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   1000
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   "Fecha|Evento"
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
   Begin VB.Menu MnuEventos 
      Caption         =   "Eventos"
      Visible         =   0   'False
      Begin VB.Menu MnuEvAddMenu 
         Caption         =   "Agregar evento"
         Begin VB.Menu MnuEvAdd 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu MnuEvDel 
         Caption         =   "Eliminar Evento"
      End
   End
End
Attribute VB_Name = "frmEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public prmIDServicio As Long
Private Sub s_LoadMenuEventos()
Dim rsE As rdoResultset
Dim iQ As Integer
'    Sólo la descripción (en caso de q en el comentario lo ocultes), o la descripción y entre parentesis la clave en caso de q en el comentario NO lo ocultes.
    
    If MnuEvAdd(0).Caption <> "" Then Exit Sub
    iQ = 0
    Cons = "Select rTrim(ESeClave), rtrim(ESeDescripcion) From EventosServicio Order By ESeOrden"
    Set rsE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsE.EOF
        If iQ > 0 Then
            Load MnuEvAdd(iQ)
        End If
        With MnuEvAdd(iQ)
            .Visible = True
            .Enabled = True
            .Caption = rsE(1)
            .Tag = rsE(0)
        End With
        iQ = iQ + 1
        rsE.MoveNext
    Loop
    rsE.Close
End Sub

Private Sub Form_Load()
    ObtengoSeteoForm Me
End Sub

Private Sub Form_Resize()
On Error Resume Next
    With vsGrid
        .Move 0, 0, ScaleWidth, ScaleHeight
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 1, , False
    End With
End Sub

Public Sub s_Clean()
    vsGrid.Rows = 1
End Sub
Public Sub s_FillGrid()
On Error GoTo errFG
Dim rsFG As rdoResultset
Dim sAux As String
Dim sLR As String
Dim sLI As String

    s_LoadMenuEventos
    vsGrid.Rows = 1
    vsGrid.ColDataType(0) = flexDTDate
    If prmIDServicio = 0 Then Exit Sub
    
    Cons = "Select SerFecha, SerComentario, rTrim(UsuIdentificacion) as Usu, VisTipo, VisFecha, VisSinEfecto, VisComentario, CamNombre, rTrim(LocNombre) as LIn, SucAbreviacion From Servicio " _
        & " Left Outer Join ServicioVisita On VisServicio = SerCodigo " _
        & " Left Outer Join Camion On VisCamion = CamCodigo " _
        & " Left Outer Join Sucursal On SerLocalReparacion = SucCodigo " _
        & ", Usuario, Local Where SerCodigo = " & prmIDServicio & " And SerUsuario = UsuCodigo And SerLocalIngreso = LocCodigo"
    
    Set rsFG = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsFG.EOF Then
        
        If Not IsNull(rsFG!SucAbreviacion) Then sLR = Trim(rsFG!SucAbreviacion)
        If Not IsNull(rsFG!LIn) Then sLI = Trim(rsFG!LIn)
        
        With vsGrid
            .ColWidth(1) = 1000
            .ColFormat(0) = "dd/mm/yy"

            .ColAlignment(1) = flexAlignLeftTop
            .AddItem rsFG!SerFecha  'Format(rsFG!SerFecha, "dd/mm/yy")
            .Cell(flexcpText, .Rows - 1, 1) = "Ingresado por " & rsFG!Usu & " en " & rsFG!LIn
                
            Do While Not rsFG.EOF
                
                If Not IsNull(rsFG!VisTipo) Then
                    .AddItem rsFG!VisFecha          'Format(rsFG!VisFecha, "dd/mm/yy")
                    
                    sAux = IIf(rsFG!VisTipo = TipoServicio.Entrega, "Entrega", IIf(rsFG!VisTipo = TipoServicio.Retiro, "Retiro", "Visita"))
                    If Not rsFG!VisSinEfecto Then
                        sAux = "Se ANULO " & sAux
                    End If
                    sAux = sAux & ", camión: " & Trim(rsFG!CamNombre)
                    If Not IsNull(rsFG!VisComentario) Then
                        sAux = sAux & "; comentario: " & Trim(rsFG!VisComentario)
                    End If
                    .Cell(flexcpText, .Rows - 1, 1) = sAux
                End If
                
                rsFG.MoveNext
            Loop
            rsFG.Close
            
            Cons = "Select Taller.*, CamTIda.CamNombre as CIda, CamTVuelta.CamNombre as CVuelta From Taller " & _
                            " Left Outer Join Camion as CamTIda on CamTIda.CamCodigo = TalIngresoCamion " & _
                            " Left Outer Join Camion as CamTVuelta on CamTVuelta.CamCodigo = TalSalidaCamion " & _
                    " Where TalServicio = " & prmIDServicio
            Set rsFG = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rsFG.EOF Then
            
                If Not IsNull(rsFG!TalIngresoCamion) Then
                    'Tengo un traslado de ingreso.
                    '------------------------------------------------------
                    .AddItem rsFG!TalFIngresoRealizado  'Format(rsFG!TalFIngresoRealizado, "dd/mm/yy")
                    .Cell(flexcpText, .Rows - 1, 1) = "Trasladado IDA por: " & Trim(rsFG!CIda) & " a " & sLR
                    
                    If Not IsNull(rsFG!TalFIngresoRecepcion) Then
                        .AddItem Format(rsFG!TalFIngresoRecepcion, "dd/mm/yy")
                        .Cell(flexcpText, .Rows - 1, 1) = "Ingresó al local de reparación (fin traslado)."
                    End If
                End If
                
                If Not IsNull(rsFG!TalSalidaCamion) Then
                    'Tengo un traslado de ingreso.
                    '------------------------------------------------------
                    .AddItem rsFG!TalFSalidaRealizado   'Format(rsFG!TalFSalidaRealizado, "dd/mm/yy")
                    .Cell(flexcpText, .Rows - 1, 1) = "Trasladado Vuelta por: " & Trim(rsFG!CVuelta) & " a  origen " & sLI
                    
                    If Not IsNull(rsFG!TalFSalidaRecepcion) Then
                        .AddItem rsFG!TalFSalidaRecepcion   'Format(rsFG!TalFSalidaRecepcion, "dd/mm/yy")
                        .Cell(flexcpText, .Rows - 1, 1) = "Recepción en local de entrega."
                    End If
                End If
                
                If Not IsNull(rsFG!TalFPresupuesto) Then
                    .AddItem rsFG!TalFPresupuesto   'Format(rsFG!TalFPresupuesto, "dd/mm/yy")
                    .Cell(flexcpText, .Rows - 1, 1) = "Se presupuestó."
                End If
                
                If Not IsNull(rsFG!TalFAceptacion) Then
                    .AddItem rsFG!TalFAceptacion        'Format(rsFG!TalFAceptacion, "dd/mm/yy")
                    If rsFG!TalAceptado Then
                        sAux = "Cliente acepto el presupuesto."
                    Else
                        sAux = "Cliente NO acepto el presupuesto."
                    End If
                    .Cell(flexcpText, .Rows - 1, 1) = sAux
                End If
                
                If Not IsNull(rsFG!TalFReparado) Then
                    .AddItem rsFG!TalFReparado  'Format(rsFG!TalFReparado, "dd/mm/yy")
                    If rsFG!TalSinArreglo Then
                        sAux = "Se indico que el producto no tiene arreglo."
                    Else
                        sAux = "Se dio por reparado el producto"
                    End If
                    .Cell(flexcpText, .Rows - 1, 1) = sAux
                End If
                
                If Not IsNull(rsFG!TalComentario) Then
                    sAux = rsFG!TalComentario
                    'Busco los eventos de taller.
                    'Las claves las guardo como [idKey:Fecha;idKey2:Fecha;....]
                    sAux = f_GetEventos(sAux)
                    If sAux <> "" Then s_AddEventos Trim(sAux)
                End If
            End If
            'CIERRO abajo
        
            'ESTO ES LO QUE HACE QUE LA ROW SE AGRANDE
            .WordWrap = True
            .AutoSizeMode = flexAutoSizeRowHeight
            .AutoSize 1, , False
                        
            .Select 1, 0, .Rows - 1, 0
            .Sort = flexSortGenericAscending
            .Select 1, 0
        End With
    End If
    rsFG.Close
    Exit Sub
errFG:
    clsGeneral.OcurrioError "Error al cargar los datos del servicio.", Err.Description
End Sub

Private Sub s_AddEventos(ByVal sEventos As String)
Dim vEventos() As String
Dim vClave() As String
Dim iQ As Integer
Dim rsKey As rdoResultset

    On Error GoTo errAE

    sEventos = Replace(Replace(sEventos, "]", ""), "[", "")
    If sEventos = "" Then Exit Sub
    If InStr(sEventos, ":") = 0 Then Exit Sub
    
    vEventos = Split(sEventos, ";")
    
    For iQ = 0 To UBound(vEventos)
        If vEventos(iQ) <> "" Then
            vClave = Split(vEventos(iQ), ":")
            If Trim(vClave(0)) <> "" Then
                Cons = "Select rTrim(EseDescripcion) From EventosServicio Where ESeClave = '" & vClave(0) & "'"
                Set rsKey = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not rsKey.EOF Then
                    vsGrid.AddItem vClave(1)
                    vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 1) = rsKey(0)
                    vsGrid.Cell(flexcpData, vsGrid.Rows - 1, 0) = vEventos(iQ)
                End If
                rsKey.Close
            End If
        End If
    Next iQ
    Exit Sub
errAE:
    clsGeneral.OcurrioError "Error al cargar los eventos para el servicio.", Err.Description
End Sub

Private Sub DeleteEvento()
Dim rsEv As rdoResultset
Dim sAux As String, sMemo As String
    On Error GoTo errDE
    Cons = "Select * From Taller where TalServicio = " & prmIDServicio
    Set rsEv = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsEv.EOF Then
        If Not IsNull(rsEv!TalComentario) Then
            sAux = f_GetEventos(Trim(rsEv!TalComentario))
            sMemo = Replace(Trim(rsEv!TalComentario), sAux, "")
            If sAux <> "" Then
                sAux = Replace(Trim(sAux), vsGrid.Cell(flexcpData, vsGrid.Row, 0), "", , 1, vbTextCompare)
                sAux = Replace(Replace(sAux, "[", ""), "]", "")
                sAux = Trim(sAux)
                If sAux <> "" Then
                    If Mid(sAux, 1, 1) = ";" Then sAux = Mid(sAux, 2)
                    sAux = Replace(sAux, ";;", ";")
                    sMemo = "[" & sAux & "]" & sMemo
                End If
            End If
            rsEv.Edit
            rsEv!TalComentario = IIf(sMemo <> "", sMemo, Null)
            rsEv.Update
        End If
    End If
    rsEv.Close
    On Error Resume Next
    s_FillGrid
    Exit Sub
errDE:
    clsGeneral.OcurrioError "Error al eliminar el evento.", Err.Description
End Sub


Private Sub Form_Unload(Cancel As Integer)
    GuardoSeteoForm Me
End Sub

Private Sub MnuEvAdd_Click(Index As Integer)
    AddEvento MnuEvAdd(Index).Tag, prmIDServicio
    On Error Resume Next
    s_FillGrid
End Sub

Private Sub MnuEvDel_Click()
    DeleteEvento
End Sub

Private Sub vsGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 93 And prmIDServicio > 0 Then
        MnuEvDel.Enabled = False
        If vsGrid.Row >= vsGrid.FixedRows Then
            If vsGrid.Cell(flexcpData, vsGrid.Row, 0) <> "" Then MnuEvDel.Enabled = True
        End If
        PopupMenu MnuEventos, , vsGrid.ColPos(0), vsGrid.RowPos(vsGrid.Row)
    End If
End Sub

Private Sub vsGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And prmIDServicio > 0 Then
        MnuEvDel.Enabled = False
        If vsGrid.Row >= vsGrid.FixedRows Then
            If vsGrid.Cell(flexcpData, vsGrid.Row, 0) <> "" Then MnuEvDel.Enabled = True
        End If
        PopupMenu MnuEventos
    End If
End Sub

