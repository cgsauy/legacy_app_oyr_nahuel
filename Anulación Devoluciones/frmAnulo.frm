VERSION 5.00
Begin VB.Form frmAnulo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anulación de Devolución"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAnulo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bCancel 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4440
      TabIndex        =   21
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton bOK 
      Caption         =   "&Anular"
      Height          =   375
      Left            =   3360
      TabIndex        =   20
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox tCodigo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      MaxLength       =   7
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lAccion 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   23
      Top             =   3720
      Width           =   5055
   End
   Begin VB.Label lAnulada 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Anulada 15-12-2004 15:58"
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
      Height          =   255
      Left            =   2040
      TabIndex        =   22
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lEnvio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Local:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   19
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Envío:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Recepción:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   17
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comentario:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Local:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Artículo:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nota:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lTipoDoc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Documento:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lTitulo 
      BackColor       =   &H8000000C&
      Caption         =   " Información de la anulación"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   5295
   End
   Begin VB.Label lAltaLocal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lArticulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Artículo:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Label lComentario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comentario:"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   1200
      TabIndex        =   7
      Top             =   2760
      Width           =   4095
   End
   Begin VB.Label lLocal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Local:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lNota 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nota:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   960
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   " Detalle Ficha"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Código:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmAnulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub f_SetMsgAccion()
    
    If bOK.Enabled Then
        If lAltaLocal.Caption <> "" Then
            'Ya ingresó la mercadería al local.
            lAccion.Caption = "Los artículos ya fueron ingresados al local, se hará movimiento de stock."
        Else
            lAccion.Caption = "Sin movimiento de stock."
        End If
    Else
        If lAnulada.Caption <> "" Then
            lAccion.Caption = "La devolución ya fue anulada."
        ElseIf Val(lNota.Tag) > 0 And lAltaLocal.Caption <> "" Then
            lAccion.Caption = "Devolución completa, tiene alta en local y tiene nota."
        End If
    End If
End Sub

Private Sub frm_Clean()

    lCliente.Caption = ""
    With lDocumento
        .Caption = "": .Tag = ""
    End With
    With lNota
        .Caption = "": .Tag = ""
    End With

    lLocal.Caption = ""
    With lArticulo
        .Caption = ""
        .Tag = ""           'Al grabar el id de artículo.
    End With
    
    lAltaLocal.Caption = ""
    lEnvio.Caption = ""
    With lComentario
        .Caption = ""
        .Tag = ""               'Lo uso al grabar para la cantidad de artículos
    End With
    lAnulada.Caption = ""
    lAccion.Caption = ""
    lTipoDoc.Tag = ""
    bOK.Enabled = False
    
End Sub

Private Sub f_Foco(C As Control)
    On Error Resume Next
    If C.Enabled Then
        C.SelStart = 0
        C.SelLength = Len(C.Text)
        C.SetFocus
    End If
End Sub

Private Sub bCancel_Click()
    Unload Me
End Sub

Private Sub bOK_Click()
Dim bStock As Boolean
Dim lUID As Long, sDef As String

    If MsgBox("¿Confirma anular la devolución?", vbQuestion + vbYesNo, "Anular Ficha de Devolución") = vbNo Then Exit Sub
    
    FechaDelServidor
    
    '............................................................Suceso
    Dim objSuceso As New clsSuceso
    With objSuceso
        .ActivoFormulario paCodigoDeUsuario, "Anulación de Ficha de Devolución", cBase
        lUID = .Usuario
        sDef = .Defensa
    End With
    Set objSuceso = Nothing
    '............................................................Suceso
    If lUID = 0 Then Exit Sub
    
    Screen.MousePointer = 11
    On Error GoTo ErrBT
    cBase.BeginTrans
    On Error GoTo ErrRB
    
    'Vuelvo a cargar la dichosa.
    Cons = "Select * From Devolucion Where DevID = " & Val(tCodigo.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        cBase.RollbackTrans
        Screen.MousePointer = 0
        MsgBox "No se encontró el registro de la devolución.", vbExclamation, "Atención"
        Exit Sub
    Else
        If Not IsNull(RsAux!DevAnulada) Then
            RsAux.Close
            cBase.RollbackTrans
            Screen.MousePointer = 0
            MsgBox "La ficha ya fue anulada.", vbExclamation, "Atención"
            Exit Sub
        Else
            If Not IsNull(RsAux!DevNota) And Not IsNull(RsAux!DevFAltalocal) Then
                RsAux.Close
                cBase.RollbackTrans
                Screen.MousePointer = 0
                MsgBox "La ficha fue completada, verifique.", vbExclamation, "Atención"
                Exit Sub
            Else
                'Vamo arriba.
                If Not IsNull(RsAux!DevFAltalocal) Then
                    bStock = True
                    lLocal.Tag = RsAux!DevLocal
                Else
                    bStock = False
                End If
                
                'Para retornar los artículos al documento.
                If Not IsNull(RsAux!DevFactura) Then
                    If Val(lDocumento.Tag) <> RsAux!DevFactura Then
                        RsAux.Close
                        cBase.RollbackTrans
                        MsgBox "La ficha fue modificada, por favor cargue los datos nuevamente.", vbExclamation, "Atención"
                        Exit Sub
                    End If
                    lDocumento.Tag = RsAux!DevFactura
                Else
                    lDocumento.Tag = ""
                End If
                
                lArticulo.Tag = RsAux!DevArticulo
                lComentario.Tag = RsAux!DevCantidad
                
                RsAux.Edit
                RsAux!DevAnulada = Format(Now, "yyyy/mm/dd hh:nn:ss")
                RsAux.Update
                
            End If
        End If
    End If
    RsAux.Close
    
    
    If bStock Then
    'Hago alta en stock del local y del stock total.
        MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, Val(lLocal.Tag), Val(lArticulo.Tag), Val(lComentario.Tag), paEstadoArticuloEntrega, -1
        MarcoMovimientoStockTotal Val(lArticulo.Tag), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, Val(lComentario.Tag), -1
        'Si hay documento hago los movimientos con el documento.
        If Val(lDocumento.Tag) > 0 Then
            MarcoMovimientoStockFisico lUID, TipoLocal.Deposito, Val(lLocal.Tag), Val(lArticulo.Tag), Val(lComentario.Tag), paEstadoArticuloEntrega, -1, Val(lTipoDoc.Tag), Val(lDocumento.Tag)
        Else
            MarcoMovimientoStockFisico lUID, TipoLocal.Deposito, Val(lLocal.Tag), Val(lArticulo.Tag), Val(lComentario.Tag), paEstadoArticuloEntrega, -1, 28, Val(tCodigo.Tag)
        End If
    End If
    
    'tiposuceso varios stock = 98
    clsGeneral.RegistroSuceso cBase, Now, 98, paCodigoDeTerminal, lUID, Val(lDocumento.Tag), Val(lArticulo.Tag), "Anulación Ficha devolución: " & Val(tCodigo.Tag), sDef
    
    cBase.CommitTrans
    tCodigo.Text = ""
    Screen.MousePointer = 0
    Exit Sub
    
ErrBT:
    clsGeneral.OcurrioError "Ocurrio un error al intentar iniciar la transacción.", Err.Description
    Screen.MousePointer = 0

ErrVA:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error al almacenar la información.", Err.Description
    Exit Sub

ErrRB:
    Resume ErrVA
    Exit Sub
    
End Sub

Private Sub Form_Load()
    ObtengoSeteoForm Me, 400, 400
    frm_Clean
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    GuardoSeteoForm Me
    CierroConexion
    Set clsGeneral = Nothing
End Sub

Private Sub Label1_Click()
    f_Foco tCodigo
End Sub

Private Sub tCodigo_Change()
    If Val(tCodigo.Tag) > 0 Then
        tCodigo.Tag = ""
        frm_Clean
    End If
End Sub

Private Sub tCodigo_GotFocus()
    f_Foco tCodigo
End Sub

Private Sub tCodigo_KeyPress(KeyAscii As Integer)
On Error GoTo errLD
    If KeyAscii = vbKeyReturn Then
        If Val(tCodigo.Tag) = 0 Then
            'Busco.
            If Not IsNumeric(tCodigo.Text) Then Exit Sub
            
            frm_Clean
            Cons = "Select Devolucion.*, CliTipo, CliCIRuc, CPeNombre1, CPeApellido1, CEmFantasia, CEmNombre," & _
                        " DF.DocTipo as DFTipo, DF.DocSerie as DFSerie, DF.DocNumero as DFNro, DN.DocTipo as DNTipo, DN.DocSerie as DNSerie, DN.DocNumero as DNNro, " & _
                        " ArtCodigo, ArtNombre, IsNull(LocNombre, '') as LocNombre From Devolucion" & _
                            " Left Outer Join Cliente On DevCliente = CliCodigo" & _
                                " Left Outer Join CPersona On CliCodigo = CPeCliente" & _
                                " Left Outer Join CEmpresa On CliCodigo = CEmCliente" & _
                            " Left Outer Join Documento DF On DevFactura = DF.DocCodigo" & _
                            " Left Outer Join Documento DN On DevNota = DN.DocCodigo" & _
                            " Left Outer Join Local On DevLocal = LocCodigo" & _
                        ", Articulo" & _
                        " Where DevID = " & Val(tCodigo.Text) & " And DevArticulo = ArtID"

            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                '..............................................................................................cliente
                lCliente.Caption = ""
                If RsAux!CliTipo = 1 Then
                    lCliente.Caption = Trim(RsAux!CPeNombre1) & " " & Trim(RsAux!CPeApellido1)
                    If Not IsNull(RsAux!CliCIRuc) Then lCliente.Caption = "(" & clsGeneral.RetornoFormatoCedula(RsAux!CliCIRuc) & ") " & lCliente.Caption
                Else
                    If Not IsNull(RsAux!CEmFantasia) Then
                        lCliente.Caption = Trim(RsAux!CEmFantasia)
                    Else
                        lCliente.Caption = Trim(RsAux!CEmNombre)
                    End If
                    If Not IsNull(RsAux!CliCIRuc) Then lCliente.Caption = "(" & clsGeneral.RetornoFormatoRuc(RsAux!CliCIRuc) & ") " & lCliente.Caption
                End If
                '..............................................................................................cliente
                
                If Not IsNull(RsAux!DevFactura) Then
                    lDocumento.Tag = RsAux!DevFactura
                    lTipoDoc.Tag = RsAux!DFTipo
                    If RsAux!DFTipo = 1 Then
                        lDocumento.Caption = "Ctdo. " & Trim(RsAux!DFSerie) & " " & Trim(RsAux!DFNro)
                    Else
                        lDocumento.Caption = "Créd. " & Trim(RsAux!DFSerie) & " " & Trim(RsAux!DFNro)
                    End If
                End If
                If Not IsNull(RsAux!DevNota) Then
                    lNota.Tag = RsAux!DevNota
                    If RsAux!DNTipo = 4 Then
                        lNota.Caption = "N Créd. "
                    ElseIf RsAux!DNTipo = 3 Then
                        lNota.Caption = "N Ctdo. "
                    Else
                        lNota.Caption = lNota.Caption & "N Esp. "
                    End If
                    lNota.Caption = lNota.Caption & Trim(RsAux!DNSerie) & " " & Trim(RsAux!DNNro)
                End If
                
                lArticulo.Caption = RsAux!DevCantidad & " *  (" & RsAux!ArtCodigo & ") " & Trim(RsAux!ArtNombre)
                lLocal.Caption = Trim(RsAux!LocNombre)
                If Not IsNull(RsAux!DevFAltalocal) Then lAltaLocal.Caption = Format(RsAux!DevFAltalocal, "dd/mm/yyyy")
                If Not IsNull(RsAux!DevComentario) Then lComentario.Caption = Trim(RsAux!DevComentario)
                If Not IsNull(RsAux!DevEnvio) Then lEnvio.Caption = Trim(RsAux!DevEnvio)
                If Not IsNull(RsAux!DevAnulada) Then lAnulada.Caption = "Anulada el " & Format(RsAux!DevAnulada, "dd-mm-yyyy  hh:nn")
                
                tCodigo.Tag = tCodigo.Text
            Else
                MsgBox "No se encontró una ficha con el código ingresado.", vbInformation, "Atención"
            End If
            RsAux.Close
            f_Foco tCodigo
            
            'Condiciones para que una devolución este todavía incompleta.
            ' no anulada
            ' que tenga nota y/o factura y no tenga local
            ' tenga local y no tenga nota.
            bOK.Enabled = lAnulada.Caption = "" And ((Val(lNota.Tag) > 0 And lAltaLocal.Caption = "") Or (lAltaLocal.Caption <> "" And Val(lNota.Tag) = 0))
            
            'Doy mensaje de acción
            f_SetMsgAccion
            
        Else
            If bOK.Enabled Then bOK.SetFocus
        End If
    End If
    Exit Sub
errLD:
    clsGeneral.OcurrioError "Error al buscar la devolución.", Err.Description
End Sub

Private Sub VueltaAtrasAnulacion()
Dim bStock As Boolean
Dim lUID As Long, sDef As String

    If MsgBox("¿Confirma desanular la devolución?", vbQuestion + vbYesNo, "Anular Ficha de Devolución") = vbNo Then Exit Sub
    
    FechaDelServidor
    
    Screen.MousePointer = 11
    On Error GoTo ErrBT
    cBase.BeginTrans
    On Error GoTo ErrRB
    
    'Vuelvo a cargar la dichosa.
    Cons = "Select * From Devolucion Where DevID = " & Val(tCodigo.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        cBase.RollbackTrans
        Screen.MousePointer = 0
        MsgBox "No se encontró el registro de la devolución.", vbExclamation, "Atención"
        Exit Sub
    Else
        If IsNull(RsAux!DevAnulada) Then
            RsAux.Close
            cBase.RollbackTrans
            Screen.MousePointer = 0
            MsgBox "La ficha ya fue desanulada.", vbExclamation, "Atención"
            Exit Sub
        Else
            'Vamo arriba.
            If Not IsNull(RsAux!DevFAltalocal) Then
                bStock = True
                lLocal.Tag = RsAux!DevLocal
            Else
                bStock = False
            End If
            
            'Para retornar los artículos al documento.
            If Not IsNull(RsAux!DevFactura) Then
'                If Val(lDocumento.Tag) <> RsAux!DevFactura Then
'                    RsAux.Close
'                    cBase.RollbackTrans
'                    MsgBox "La ficha fue modificada, por favor cargue los datos nuevamente.", vbExclamation, "Atención"
'                    Exit Sub
'                End If
                lDocumento.Tag = RsAux!DevFactura
            Else
                lDocumento.Tag = ""
            End If
            
            lArticulo.Tag = RsAux!DevArticulo
            lComentario.Tag = RsAux!DevCantidad
            
            RsAux.Edit
            RsAux!DevAnulada = Null
            RsAux.Update

        End If
    End If
    RsAux.Close
    
    
    If bStock Then
    'Hago alta en stock del local y del stock total.
        MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, Val(lLocal.Tag), Val(lArticulo.Tag), Val(lComentario.Tag), paEstadoArticuloEntrega, 1
        MarcoMovimientoStockTotal Val(lArticulo.Tag), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, Val(lComentario.Tag), 1
    
        'Si hay documento hago los movimientos con el documento.
        If Val(lDocumento.Tag) > 0 Then
            MarcoMovimientoStockFisico lUID, TipoLocal.Deposito, Val(lLocal.Tag), Val(lArticulo.Tag), Val(lComentario.Tag), paEstadoArticuloEntrega, 1, Val(lTipoDoc.Tag), Val(lDocumento.Tag)
        Else
            MarcoMovimientoStockFisico lUID, TipoLocal.Deposito, Val(lLocal.Tag), Val(lArticulo.Tag), Val(lComentario.Tag), paEstadoArticuloEntrega, 1, 28, Val(tCodigo.Tag)
        End If
    End If
    
    'tiposuceso varios stock = 98
'    clsGeneral.RegistroSuceso cBase, Now, 98, paCodigoDeTerminal, lUID, Val(lDocumento.Tag), Val(lArticulo.Tag), "Anulación Ficha devolución: " & Val(tCodigo.Tag), sDef
    
    cBase.CommitTrans
    tCodigo.Text = ""
    Screen.MousePointer = 0
    Exit Sub
    
ErrBT:
    clsGeneral.OcurrioError "Ocurrio un error al intentar iniciar la transacción.", Err.Description
    Screen.MousePointer = 0

ErrVA:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error al almacenar la información.", Err.Description
    Exit Sub

ErrRB:
    Resume ErrVA
    Exit Sub

End Sub
