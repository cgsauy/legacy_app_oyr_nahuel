VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmFac 
   BackColor       =   &H80000006&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11115
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   14730
   ControlBox      =   0   'False
   FillColor       =   &H00400000&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFactura.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11115
   ScaleWidth      =   14730
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser wbNav 
      Height          =   2535
      Left            =   480
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   840
      Width           =   5535
      ExtentX         =   9763
      ExtentY         =   4471
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   360
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFactura.frx":17D2A
            Key             =   "exc"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFactura.frx":2FA64
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFactura.frx":4779E
            Key             =   "inf"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFactura.frx":5F4D8
            Key             =   "car"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFactura.frx":77212
            Key             =   "sto"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox tcBarra 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12120
      TabIndex        =   0
      Top             =   10440
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   10080
      Top             =   360
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsGridArt 
      Height          =   1935
      Left            =   960
      TabIndex        =   1
      Top             =   2280
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   3413
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483630
      ForeColor       =   -2147483634
      BackColorFixed  =   -2147483630
      ForeColorFixed  =   -2147483634
      BackColorSel    =   14599344
      ForeColorSel    =   6291456
      BackColorBkg    =   -2147483637
      BackColorAlternate=   -2147483630
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   0
      SheetBorder     =   0
      FocusRect       =   0
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
   Begin VSFlex6DAOCtl.vsFlexGrid vsGrid2 
      Height          =   1695
      Left            =   1080
      TabIndex        =   2
      Top             =   4800
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2990
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483630
      ForeColor       =   -2147483634
      BackColorFixed  =   -2147483630
      ForeColorFixed  =   -2147483634
      BackColorSel    =   -2147483630
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483638
      BackColorAlternate=   -2147483630
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
   Begin VB.Image imgCancelar 
      Height          =   1305
      Left            =   7440
      Picture         =   "frmFactura.frx":77AE1
      Top             =   9600
      Width           =   2850
   End
   Begin VB.Image imgNo 
      Height          =   1305
      Left            =   4320
      Picture         =   "frmFactura.frx":78399
      Top             =   9600
      Width           =   2850
   End
   Begin VB.Image imgRetirar 
      Height          =   1305
      Left            =   1200
      Picture         =   "frmFactura.frx":78CCB
      Top             =   9600
      Width           =   2850
   End
   Begin VB.Label lMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lAvisos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   1935
      Left            =   2400
      TabIndex        =   4
      Top             =   7440
      Width           =   11535
   End
   Begin VB.Image imgIcon 
      Height          =   720
      Left            =   1440
      Picture         =   "frmFactura.frx":79749
      Top             =   7440
      Width           =   720
   End
   Begin VB.Shape shMsg 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2175
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   7320
      Width           =   12975
   End
   Begin VB.Label linfoGrid 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "(Flechas) cambia de fila            (+) Agrega         (-) quita artículos        Escape = Cancela"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   4200
      Width           =   12375
   End
   Begin VB.Label lTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pronto para ingresar una nueva factura"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   735
      Left            =   960
      TabIndex        =   5
      Top             =   120
      Width           =   12135
   End
   Begin VB.Label lSerie 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   960
      TabIndex        =   3
      Top             =   1440
      Width           =   12975
   End
   Begin VB.Shape shfac 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00CEBAB3&
      FillColor       =   &H00A88D7B&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   12975
   End
End
Attribute VB_Name = "frmFac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private sWav As String

Private iHoraAlternativa As Long, iLocalAlternativo As Integer
Private sLocalAlternativo As String

Private iTipoDoc As Byte
Private iDocumento As Long
Private gFechaDocumento As Date

Private Cons As String
Private RsAux As rdoResultset

Private bCercano As Boolean 'si esta para arrimar es true

Private Enum EstadoS
    Anulado = 0
    Visita = 1
    Retiro = 2
    Taller = 3
    Entrega = 4
    Cumplido = 5
End Enum

Private Enum EstadoEntregaAuxiliar 'estado de entrega
    SinEntregar = 1
    Arrimado = 2
    Entregado = 3
    ClienteSeFue = 4
    EraParaEnvio = 5
    OtroLocal = 6
End Enum

Private Enum TipoDocumento
    'Documentos Facturacion
    Servicio = 0
    Contado = 1
    Credito = 2
    NotaDevolucion = 3
    NotaCredito = 4
    ReciboDePago = 5
    Remito = 6
End Enum

Private Enum eEstMsg
    Informo = 0
    Advierto = 1
    Pregunto = 2
    Error = 3
    Pronto = 4
End Enum

Private Sub loc_SetSonido(ByVal sFile As String)
    On Error Resume Next
    Dim Result As Long
    Result = sndPlaySound(sWav & sFile, 1)
End Sub

Private Function DoBusy()
    DoEvents
    Do While wbNav.Busy
        DoEvents
    Loop
End Function

Private Sub loc_VerArrimar()

    Cons = "Select ParValor from Parametro Where ParNombre = 'dep_Estado_Arrimar_" & paCodigoDeSucursal & "' And ParValor Is Not Null"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then paArrimar = RsAux("ParValor")
    RsAux.Close
    
End Sub

Private Function fnc_LoadServicio() As Boolean
On Error GoTo errLS
Dim sQy As String
Dim RsAux As rdoResultset
Dim iCosto As Currency

    
    sQy = "Select * From Servicio " & _
            "Left Outer Join Taller ON TalServicio = SerCodigo " & _
            "Left Outer Join Documento On SerDocumento = DocCodigo " & _
        "Where SerCodigo = " & Mid(tcBarra.Text, 2)
    Set RsAux = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    
    fnc_LoadServicio = (Not RsAux.EOF)
    
    If Not RsAux.EOF Then
        'Valido el estado
        Select Case RsAux("SerEstadoServicio")
            Case EstadoS.Taller
                If IsNull(RsAux("TalFReparado")) Then
                    loc_ShowMsg "El servicio está sin reparar o sin presupuesto aceptado, consulte en mostrador.", 9000, Informo
                Else
                    If Not IsNull(RsAux("SerCostoFinal")) Then iCosto = RsAux("SerCostoFinal") Else iCosto = -1
                
                    If IsNull(RsAux("SerDocumento")) Then
                        If iCosto > 0 Then
                            loc_ShowMsg "El servicio tiene un costo de $ " & Format(iCosto, "#,##0.00") & " y aún no fue facturado, pase por Colonia para realizar el pago del mismo.", 9000, Informo
                        ElseIf iCosto = -1 Then
                            loc_ShowMsg "El servicio aún no fue validado, consulte en mostrador.", 9000, Informo
                        Else
                            sQy = fnc_HoraArtAEntregar(0, RsAux("SerCodigo"), 0)
                            If Val(sQy) > 59 Then
                                loc_ShowMsg "Su solicitud está siendo procesada, estimamos que en " & Val(sQy) \ 60 & " minuto(s) se cumplirá la misma.", 9000, Informo
                            Else
                                'Ahora si esta con tipo de servicio no más
                                'Grabo el servicio.
                                iTipoDoc = 0
                                iDocumento = RsAux("SerCodigo")
                                loc_SaveServicio RsAux("SerCodigo")
                            End If
                        End If
                    Else
                        'OK acá estoy en condiciones
                        'Ahora valido que no este por el documento asignado.
                        sQy = fnc_HoraArtAEntregar(RsAux("DocTipo"), RsAux("DocCodigo"), 0)
                        iCosto = fnc_HoraArtAEntregar(0, RsAux("SerCodigo"), 0)
                        If Val(sQy) > 59 Or iCosto > 0 Then
                            loc_ShowMsg "Su solicitud está siendo procesada, estimamos que en " & Val(sQy) \ 60 & " minuto(s) se cumplirá la misma.", 9000, Informo
                        Else
                            'Ahora si esta con tipo de servicio no más
                            'Grabo el servicio.
                            iTipoDoc = 0
                            iDocumento = RsAux("SerCodigo")
                            loc_SaveServicio RsAux("SerCodigo")
                        End If
                    End If
                End If
                
            Case EstadoS.Entrega:
                loc_ShowMsg "El servicio " & RsAux("SerCodigo") & ", está pendiente de entrega en domicilio, consulte en mostrador.", 9000, Advierto
                
            Case Else
                
                loc_ShowMsg "Servicio " & RsAux("SerCodigo") & ", consulte en mostrador por el estado de su servicio.", 9000, Advierto
        End Select
    End If
    RsAux.Close
    Exit Function
errLS:
    loc_ShowMsg "Error al obtener la información del servicio, reintente." & vbCrLf & Err.Description, 1500, Error
End Function

Private Sub loc_HideBotones()
    imgRetirar.Visible = False
    imgNo.Visible = False
    imgCancelar.Visible = False
End Sub

Private Sub loc_SaveServicio(ByVal iServicio As Long)
On Error GoTo errSS
    
    loc_VerArrimar
    
    FechaDelServidor
    
    Cons = "Select * From EntregaAuxiliar Where EAuDocumento = " & iDocumento & " And EAuTipo = " & iTipoDoc
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)

    RsAux.AddNew
    RsAux("EAuTipo") = iTipoDoc
    RsAux("EAuDocumento") = iDocumento
    RsAux("EAuArticulo") = 0
    RsAux("EAuFechaHora") = gFechaServidor
    RsAux("EAuCantidad") = 1
    RsAux("EAuLocal") = paCodigoDeSucursal
    If paArrimar = 1 Then
        RsAux("EAuEstado") = 1
        RsAux("EAuSinArrimar") = 1
    Else
        RsAux("EAuEstado") = 2
        RsAux("EAuSinArrimar") = 0
    End If
    RsAux("EAuTiempoProm") = 60
    RsAux("EAuTiempoTotal") = 0
    RsAux.Update
    RsAux.Close

    loc_ShowMsg "Factura asociada al servicio " & iServicio & ", a la mayor brevedad será llamado desde el mostrador", 8000, Pronto
    If paWavOK <> "" Then loc_SetSonido paWavOK
    Exit Sub
    
errSS:
    loc_ShowMsg "Error al asignar el servicio.", 3000, Error
End Sub

Private Sub loc_limpiarControles()
On Error Resume Next
    
    With vsGridArt
        .Visible = False: .Rows = 0
    End With
    
    With vsGrid2
        .Visible = False: .Rows = 0
    End With
   
    lSerie.Caption = ""
    
    loc_HideBotones
    
    linfoGrid.Visible = False
    shfac.Visible = False
    
    lSerie.Visible = False
    shfac.Visible = False
    
    lTitle.Visible = True
    wbNav.Visible = (paPagHtml <> "")
    
    loc_HideMsg
    Foco tcBarra
    
End Sub

Private Sub Foco(C As Control)
    On Error Resume Next
    If C.Enabled Then
        C.SelStart = 0
        C.SelLength = Len(C.Text)
        C.SetFocus
    End If
End Sub

Private Sub loc_CargoDatosSucursal()
On Error GoTo errCDS
Dim sQy As String
Dim rsA As rdoResultset
    
    iHoraAlternativa = 0
    sQy = "Select LocEntregaEmergencia, LocLocalAlternativo From Local where LocCodigo = " & paCodigoDeSucursal
    Set rsA = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    If Not rsA.EOF Then
        If Not IsNull(rsA("LocEntregaEmergencia")) Then iHoraAlternativa = rsA("LocEntregaEmergencia")
        If Not IsNull(rsA("LocLocalAlternativo")) Then iLocalAlternativo = Trim(rsA("LocLocalAlternativo"))
    End If
    rsA.Close
    
    If iLocalAlternativo > 0 Then
        sQy = "Select LocNombre from Local Where LocCodigo = " & iLocalAlternativo
        Set rsA = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
        If Not rsA.EOF Then sLocalAlternativo = Trim(rsA("LocNombre"))
        rsA.Close
    End If
    Exit Sub
errCDS:
End Sub
Private Function fnc_GetTimeEntrega(ByVal dFechaHora As Date, ByVal iTimeProm As Long) As Long
On Error GoTo errAEntregar
Dim iTime As Long
    fnc_GetTimeEntrega = 60
    iTime = Abs(DateDiff("s", dFechaHora, Now))
    If iTime < iTimeProm Then
        If iTimeProm - iTime > 60 Then
            fnc_GetTimeEntrega = iTimeProm - iTime
        End If
    End If
Exit Function
errAEntregar:
End Function

Private Function fnc_HoraArtAEntregar(ByVal iTDoc As Byte, ByVal iDoc As Long, ByVal idArt As Long) As Long
On Error GoTo errAEntregar
Dim sCons As String
Dim RsAux As rdoResultset
Dim iTime As Long

    fnc_HoraArtAEntregar = 0
    sCons = " Select EAUFechaHora, EAuTiempoProm From EntregaAuxiliar " _
          & "Where EAuTipo = " & iTDoc & " And EAuDocumento = " & iDoc _
          & " And EAuEstado In (1,2) And EAuArticulo = " & idArt

    Set RsAux = cBase.OpenResultset(sCons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        iTime = Abs(DateDiff("s", RsAux("EAUFechaHora"), Now))
        fnc_HoraArtAEntregar = 60
        If iTime < RsAux("EAuTiempoProm") Then
            If RsAux("EAUTiempoProm") - iTime > 60 Then
                fnc_HoraArtAEntregar = RsAux("EAUTiempoProm") - iTime
            End If
        End If
    End If
    RsAux.Close
    
Exit Function
errAEntregar:
    fnc_HoraArtAEntregar = -10
End Function

Private Sub loc_HideMsg()
    lMsg.Visible = False
    imgIcon.Visible = False
    shMsg.Visible = False
End Sub

Private Sub Form_Click()
On Error Resume Next
    tcBarra.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)


    If KeyCode = vbKeyF7 Then
        On Error GoTo errBDocManual
        Dim snrodoc  As String
        snrodoc = InputBox("Ingrese los últimos 5 dígitos o todo el número de la factura. (sin serie)", "Búsqueda manual de facturas recientes", "")
        tcBarra.SetFocus
        If snrodoc <> "" Then
            Screen.MousePointer = 11
            Cons = "EXEC prg_EntregaCliente_BuscoDocumento " & CLng(snrodoc) & ", 200" & ", " & paCodigoDeSucursal
            Dim tipo As Byte
            Dim codigo As Long
            Dim objHelp As New clsListadeAyuda
            If objHelp.ActivarAyuda(cBase, Cons, 7000, 2, "Búsqueda manual") > 0 Then
                codigo = objHelp.RetornoDatoSeleccionado(0)
                tipo = objHelp.RetornoDatoSeleccionado(1)
            End If
            Set objHelp = Nothing
            If codigo > 0 Then BuscoDocumento tipo, codigo
            Screen.MousePointer = 0
        End If
        
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errBDocManual:
    Screen.MousePointer = 0
    MsgBox "Error al buscar, reintente por favor.", vbCritical, "Error"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo errKP
    
    If vsGridArt.Rows = 0 And vsGrid2.Rows = 0 Then Exit Sub
    
    Select Case KeyAscii
        Case Asc(LCase(paEntregaTotal)), Asc(UCase(paEntregaTotal))
            imgRetirar_Click
            KeyAscii = 0
        
        Case Asc(LCase(paEntregaParcial)), Asc(UCase(paEntregaParcial))
            imgNo_Click
            KeyAscii = 0
        
        Case Asc(LCase(paEntregaCancelar)), Asc(UCase(paEntregaCancelar))
            imgCancelar_Click
            KeyAscii = 0
            
            
    End Select
errKP:
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyN Then tcBarra.Text = ""
End Sub

Private Sub Form_Load()
On Error Resume Next
    ObtengoSeteoForm Me, 100, 100
    
    ChDir App.Path
    ChDir ("..")
    If Dir(CurDir & "\Sonidos", vbDirectory) <> "" Then sWav = CurDir & "\Sonidos\"
    
    loc_CargoDatosSucursal
    loc_HideMsg
    
    With vsGrid2
        .ColWidth(0) = 4000
        .ColWidth(1) = 500
        .ColHidden(3) = True
        .ColHidden(2) = True
        .BackColorBkg = vbBlack
    End With
    
    With vsGridArt
        .ColHidden(2) = True
        .ColHidden(3) = True
        .BackColorBkg = vbBlack
    End With
    
    If paPagHtml <> "" Then
        wbNav.Navigate2 paPagHtml
        DoBusy
    Else
        wbNav.Visible = True
    End If
    
    
    tcBarra.Top = -2000
    imgRetirar.Visible = False
    imgNo.Visible = False
    imgCancelar.Visible = False
    loc_limpiarControles
    linfoGrid.Visible = False
    shfac.Visible = False
    tcBarra.SetFocus
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    shfac.Move 480, shfac.Top, Me.ScaleWidth - 960
    lSerie.Left = shfac.Left + 120
    lSerie.Width = shfac.Width - 240
    
    With vsGridArt
        .Left = 720
        .Width = ScaleWidth - 1440
        .ColWidth(0) = .Width - 2500
    End With
    linfoGrid.Move 720, linfoGrid.Top, vsGridArt.Width
    
    shMsg.Move 1200, Me.ScaleHeight - 4500, ScaleWidth - 2400
    imgIcon.Move shMsg.Left + 240, shMsg.Top + 240
    lMsg.Move imgIcon.Left + 240 + imgIcon.Width, imgIcon.Top, shMsg.Width - (imgIcon.Left + 360 + imgIcon.Width)
    
    imgNo.Top = Me.ScaleHeight - 400 - imgNo.Height
    imgCancelar.Top = imgNo.Top
    imgRetirar.Top = imgNo.Top
    
    lTitle.Move 0, Me.ScaleHeight - 3000, ScaleWidth
    
    wbNav.Move 0, 0, ScaleWidth, ScaleHeight
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    cBase.Close
    Set clsGeneral = Nothing
    GuardoSeteoForm Me
End Sub

Private Sub imgCancelar_Click()
    loc_limpiarControles
    loc_HideMsg
    tcBarra.Text = ""
    tcBarra.SetFocus
End Sub

Private Sub imgNo_Click()
On Error Resume Next
    loc_HideMsg
    tcBarra.Text = ""
    Timer1.Enabled = False
    If vsGridArt.Visible And vsGridArt.Enabled Then
        vsGridArt.SetFocus
        vsGridArt.Select 0, 0, 0, 1
    End If
End Sub

Private Sub imgRetirar_Click()
Dim iSgdos As Long
Dim bHay As Long
    
    tcBarra.Enabled = False
    vsGridArt.Enabled = False
    
    'freno timer
    Timer1.Enabled = False
    
    'Recorro grilla para determinar si realmente hay algo seleccionado.
    For iSgdos = 0 To vsGridArt.Rows - 1
        If Val(vsGridArt.Cell(flexcpText, iSgdos, 1)) > 0 Then bHay = True: Exit For
    Next
    
    If Not bHay Then
        loc_ShowMsg "No hay artículos seleccionados a entregar, corrija o cancele su solicitud.", 9000, Error
        tcBarra.Enabled = True
        vsGridArt.Enabled = True
        vsGridArt.SetFocus
        vsGridArt.Select 0, 0
        Exit Sub
    Else
        iSgdos = fnc_GrabarRetiro
        loc_limpiarControles
        If iSgdos > 0 Then
            iSgdos = iSgdos / 60
            loc_ShowMsg "Estimamos que en " & iSgdos & IIf(iSgdos > 1, " minutos", "minuto") & " será llamado desde el mostrador para entregarle la mercadería.", 9000, Pronto
            If paWavOK <> "" Then loc_SetSonido paWavOK
        ElseIf iSgdos = 0 Then
            loc_ShowMsg "Atención su solicitud no fue procesada debido a un error, reintente el proceso nuevamente.", 7000, Error
        End If
        'Oculto los botones
        loc_HideBotones
    End If
    tcBarra.Text = ""
    tcBarra.Enabled = True
    tcBarra.SetFocus

End Sub

Private Sub tcBarra_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn And Trim(tcBarra.Text) <> "" Then
        If LCase(tcBarra.Text) = LCase(paCloseApp) Then
            Unload Me
        Else
            FormatoBarras tcBarra.Text
        End If
    End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    Timer1.Enabled = False
    If vsGridArt.Rows > 0 And tcBarra.Enabled Then
        imgRetirar_Click
    End If
    loc_limpiarControles
End Sub

Private Sub loc_ShowMsg(ByVal sTxt As String, ByVal iInterval As Integer, ByVal iEstado As eEstMsg)
On Error Resume Next
Dim iFColor As Long
    
    shMsg.Height = 2175
    lMsg.FontSize = 24
    
    Select Case iEstado
        Case 0  'información
            imgIcon.Picture = imlIcons.ListImages("inf").Picture
            shMsg.FillColor = &HCEBAB3
            shMsg.BorderColor = &HFAF0EB
        Case 1  'advertencia
            imgIcon.Picture = imlIcons.ListImages("exc").Picture
            shMsg.FillColor = &HD0FFFF
            shMsg.BorderColor = &HC0C0&
            lMsg.ForeColor = &H40&
            lMsg.FontSize = 32
            shMsg.Height = 4000
            
        Case 2  'Pregunta
            imgIcon.Picture = imlIcons.ListImages("pre").Picture
            shMsg.FillColor = &HCEBAB3
            shMsg.BorderColor = &HFAF0EB
            
        Case 3  'Error
            imgIcon.Picture = imlIcons.ListImages("sto").Picture
            shMsg.FillColor = &HFFFFFF
            shMsg.BorderColor = &HC0&
            lMsg.ForeColor = &H40&
            lMsg.FontSize = 32
            shMsg.Height = 4000
            
        Case 4  'Fin
            'Puse este nuevo estado para subir el cuadro de msg
            shMsg.Move 1200, (Me.ScaleHeight / 2) - 1200, ScaleWidth - 2400
            imgIcon.Move shMsg.Left + 240, shMsg.Top + 240
            lMsg.Move imgIcon.Left + 240 + imgIcon.Width, imgIcon.Top, shMsg.Width - (imgIcon.Left + 360 + imgIcon.Width)
            imgIcon.Picture = imlIcons.ListImages("inf").Picture
            shMsg.FillColor = &HCEBAB3
            shMsg.BorderColor = &HFAF0EB
    End Select
    
    shMsg.Move 1200, Me.ScaleHeight - 2375 - shMsg.Height, ScaleWidth - 2400
    
    imgIcon.Move shMsg.Left + 240, shMsg.Top + 240
    lMsg.Move imgIcon.Left + 240 + imgIcon.Width, imgIcon.Top, shMsg.Width - (imgIcon.Left + 360 + imgIcon.Width), shMsg.Height
    
    lMsg.Caption = sTxt
    
    lMsg.Visible = True
    imgIcon.Visible = True
    shMsg.Visible = True
    wbNav.Visible = False
    
    Timer1.Interval = iInterval
    Timer1.Enabled = True
    
End Sub

Private Sub FormatoBarras(Texto As String)
Dim iAuxTipo As Byte
Dim iCodDoc As Long

    'Tengo un documento en espera del ok al pasar un nuevo documento -->
    If (vsGridArt.Rows > 0 And imgRetirar.Visible) Then
        imgRetirar_Click
    End If
                
    loc_HideMsg
    loc_limpiarControles

    On Error GoTo errInt
    Texto = UCase(Texto)
    If InStr(1, Texto, "D", vbTextCompare) > 0 Then
        If Not IsNumeric(Mid(Texto, 1, InStr(Texto, "D") - 1)) Then loc_ShowMsg "El dato ingresado no es un documento de compra.", 7000, Advierto: Exit Sub
        iAuxTipo = CLng(Mid(Texto, 1, InStr(Texto, "D") - 1))
        iCodDoc = CLng(Trim(Mid(Texto, InStr(Texto, "D") + 1, Len(Texto))))
    ElseIf LCase(Left(Texto, 1)) = "s" Then
        If fnc_LoadServicio Then Exit Sub
    End If
    
    Timer1.Enabled = False
    
    tcBarra.Tag = Texto
    Select Case iAuxTipo
'        Case TipoDocumento.Remito
'             BuscoRemito iCodDoc
        
        Case TipoDocumento.Contado, TipoDocumento.Credito, TipoDocumento.Remito
             BuscoDocumento iAuxTipo, iCodDoc
            
        Case Else
            loc_ShowMsg "El dato ingresado no es un documento de compra.", 7000, Advierto
            
    End Select
    Screen.MousePointer = 0
    Exit Sub

errInt:
    Screen.MousePointer = 0
    loc_ShowMsg "Error al interpretar el código de barras:" & Err.Description, 3000, Error
End Sub

'Private Sub BuscoRemito(ByVal Numero As Long)
'On Error GoTo errRemito
'    Dim iAux As Long
'
'    iTipoDoc = 0
'    iDocumento = 0
'
'    Cons = " Select NomCli = RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)+ " _
'                & " RTrim(' ' + CPeNombre2), NomEmp=(RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')'), " _
'            & " Documento.*, RemDocumento, RemCodigo, RemModificado, EAUFechaHora, EAuTiempoProm " _
'            & " From (((((Remito Inner Join Documento ON RemDocumento = DocCodigo)" _
'                    & " Left Outer Join Cliente on Documento.DocCliente = Cliente.CliCodigo)" _
'                    & " Left Outer Join Cpersona on Cliente.CliCodigo = CPersona.CpeCliente)" _
'                    & " Left Outer Join CEmpresa On CliCodigo = CEmCliente)" _
'                    & " Left Outer Join EntregaAuxiliar ON 6 = EAuTipo And RemCodigo = EAuDocumento And EAuEstado In (1,2))" _
'            & " Where RemCodigo = " & Numero
'
'
'    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
'
'    Foco tcBarra
'
'    If Not RsAux.EOF Then
'        If Not IsNull(RsAux("EAuFechaHora")) Then
'            'Está en proceso.
'            iAux = fnc_GetTimeEntrega(RsAux("EAuFechaHora"), RsAux("EAuTiempoProm"))
'            RsAux.Close
'            Screen.MousePointer = 0
'            loc_ShowMsg "Su factura está siendo procesada, estimamos que en " & iAux \ 60 & " minuto(s) se cumplirá su solicitud.", 7000, Informo
'            Foco tcBarra
'            Exit Sub
'
''13/01/2009 Un contado con remito no se puede anular por eso anulé esta sección
'
''        ElseIf RsAux!DocAnulado Then
''            Screen.MousePointer = 0
''            loc_ShowMsg "El documento ingresado está anulado, consulte en mostrador.", 7000, Advierto
''            RsAux.Close
''            Foco tcBarra
''            Exit Sub
'
''----------------------------------------------------------------------------------
'
'
'        Else
'
'            If Not IsNull(RsAux!DocPendiente) Then
'                Screen.MousePointer = 0
'                loc_ShowMsg "Documento pendiente de entrega, consulte en mostrador.", 7000, Advierto
'                RsAux.Close
'                Exit Sub
'            End If
'
'        End If
'
'    Else
'        Screen.MousePointer = 0
'        iDocumento = 0
'        loc_ShowMsg "No existe un remito para las características ingresadas, consulte en mostrador.", 7000, Advierto
'        RsAux.Close
'        Exit Sub
'    End If
'
'    iDocumento = Numero
'    iTipoDoc = TipoDocumento.Remito
'    gFechaDocumento = RsAux!RemModificado     'Siempre guardo la del Documento
'
'    Dim sNomCli As String, sFactura As String
'    If Not IsNull(RsAux!NomCli) Then sNomCli = Trim(RsAux!NomCli) Else sNomCli = RsAux("NomEmp")
'    sFactura = "Remito " & RsAux("RemCodigo")
'    RsAux.Close
'
'    If iDocumento > 0 Then
'        CargoArticulos
'        If vsGridArt.Rows > 0 Or vsGrid2.Rows > 0 Then
'            lSerie.Caption = "Remito de factura: " & sFactura & "  " & sNomCli
'            shfac.Visible = True
'            lSerie.Visible = True
'            lTitle.Visible = False
'            wbNav.Visible = False
'        End If
'    End If
'    Screen.MousePointer = 0
'    Exit Sub
'errRemito:
'    loc_ShowMsg "Error al cargar la información del remito.", 4000, Error
'End Sub
Private Sub BuscoDocumento(ByVal iTipo As Long, Optional codigo As Long = 0)
On Error GoTo errBuscoD
Dim iAux As Long, iAux1 As Long
Dim sNomCli As String, sFactura As String

    Screen.MousePointer = 11
    iDocumento = 0
    Cons = "EXEC prg_EntregaCliente_BuscoDocumento " & codigo & ", " & iTipo & ", " & paCodigoDeSucursal
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        If RsAux(0) < 0 Then
            sNomCli = Trim(RsAux(1))
            RsAux.Close
            Screen.MousePointer = 0
            loc_ShowMsg sNomCli, 7000, Advierto
            Exit Sub
        ElseIf RsAux(0) = 3 Then
            iAux = RsAux("SerCodigo")
            iDocumento = RsAux!DocCodigo
            iTipoDoc = iTipo
            gFechaDocumento = RsAux!DocFModificacion
            RsAux.Close
            'Por las dudas verifico que el servicio no este ingresado x S + IDServicio.
            iAux1 = fnc_HoraArtAEntregar(0, iAux, 0)
            If iAux1 > 0 Then
                loc_ShowMsg "Su factura de servicio está siendo procesada, estimamos que en " & iAux1 \ 60 & " minuto(s) se cumplirá su solicitud.", 7000, Informo
            Else
                loc_SaveServicio iAux
            End If
        Else
            iDocumento = RsAux!DocCodigo
            iTipoDoc = iTipo
            If Not IsNull(RsAux!DocFModificacion) Then gFechaDocumento = RsAux!DocFModificacion
            If Not IsNull(RsAux!NomCli) Then sNomCli = Trim(RsAux!NomCli)
            sFactura = RsAux("Numero")
        End If
    End If
    RsAux.Close
    If iDocumento > 0 Then
        CargoArticulos
        If vsGridArt.Rows > 0 Or vsGrid2.Rows > 0 Then
            lSerie.Caption = "Factura: " & sFactura & "  " & sNomCli
            shfac.Visible = True
            lSerie.Visible = True
            lTitle.Visible = False
            wbNav.Visible = False
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errBuscoD:
    'clsGeneral.OcurrioError "Ocurrió un error al buscar un documento.", Err.Description
End Sub

Private Function fnc_HayMovsFisicos(ByVal iArt As Long) As Boolean
Dim rsMF As rdoResultset
    
    Set rsMF = cBase.OpenResultset("Select count(*) from movimientostockfisico" & _
                        " Where MSFFecha >= GetDate()-3 And MSFLocal = " & paCodigoDeSucursal & _
                        " And MSFArticulo = " & iArt, rdOpenDynamic, rdConcurValues)
    If Not rsMF.EOF Then
        If Not IsNull(rsMF(0)) Then fnc_HayMovsFisicos = (rsMF(0) <> 0)
    End If
    rsMF.Close
                        
End Function

Private Sub CargoArticulos()
Dim rsXX As rdoResultset
Dim iAux As Long, iDocA As Long
Dim bLocalAlternativo As Boolean

    bLocalAlternativo = True
    loc_VerArrimar

    vsGridArt.Rows = 0
    vsGrid2.Rows = 0
    vsGrid2.Tag = ""

    On Error GoTo errCargar
    Cons = "EXEC prg_EntregaCliente_CargoDocumento " & iTipoDoc & ", " & iDocumento & ", " & paCodigoDeSucursal & ", " & iLocalAlternativo
    Set rsXX = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not rsXX.EOF
        If rsXX("ARetirar") > 0 Then
            If rsXX("ArtLocalRetira") = paCodigoDeSucursal Or rsXX("ArtLocalRetira") = 0 Then
                With vsGridArt
                    .AddItem Trim(rsXX("ArtNombre"))
                    .Visible = True
                    linfoGrid.Visible = True
                    
                    If rsXX("ARetirar") < rsXX("StLCantidad") Or rsXX("StLCantidad") <= 0 Then
                        .Cell(flexcpText, .Rows - 1, 1) = rsXX("ARetirar")
                    Else
                        .Cell(flexcpText, .Rows - 1, 1) = rsXX("StLCantidad")
                    End If
                    
                'DATA
                    iAux = rsXX("ArtId"): .Cell(flexcpData, .Rows - 1, 0) = iAux
                    iAux = rsXX("ARetirar"): .Cell(flexcpData, .Rows - 1, 1) = iAux
                    If Not IsNull(rsXX("AseCercano")) Then
                        .Cell(flexcpData, .Rows - 1, 3) = IIf(rsXX("AseCercano"), 2, 1)
                    Else
                        .Cell(flexcpData, .Rows - 1, 3) = 1
                    End If
                    
                End With
            Else
                vsGrid2.Visible = True
                With vsGrid2
                    .AddItem ""
                    .Cell(flexcpText, .Rows - 1, 0) = Trim(rsXX("ArtNombre"))
                    If Not IsNull(rsXX("SucAbreviacion")) Then
                        .Cell(flexcpText, .Rows - 1, 1) = "Retirar en: " & Trim(rsXX("SucAbreviacion"))
                    Else
                        .Cell(flexcpForeColor, .Rows - 1, 1) = vbRed
                        .Cell(flexcpText, .Rows - 1, 1) = "Consultar"
                    End If
                    .Tag = "otro"
                End With
            End If
            If bLocalAlternativo Then bLocalAlternativo = rsXX("LocAlternativo")
        Else
            'envIOS
            With vsGrid2
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = Trim(rsXX("ArtNombre"))
                .Cell(flexcpText, .Rows - 1, 1) = Trim(rsXX("SucAbreviacion"))
                .Tag = "otro"
            End With
        End If
        rsXX.MoveNext
    Loop
    rsXX.Close
    bLocalAlternativo = bLocalAlternativo And (vsGrid2.Rows > 0 Or vsGridArt.Rows > 0)
    
    If vsGridArt.Rows > 0 Then
'        If Not bLocalAlternativo Or sLocalAlternativo = "" Then
'            If bSinStock Then
 '               loc_ShowMsg "ATENCIÓN consulte el stock de sus artículos en el mostrador" & vbCrLf & "¿Desea retirar el/los artículo/s ?", 9000, Pregunto
  '          Else
                'loc_ShowMsg "¿Desea retirar el/los artículo/s ?", 9000, Pregunto
   '         End If
'        Else
'            If bLocalAlternativo Then
'                If vsGrid2.Rows > 0 Then
'                    loc_ShowMsg "Sugerencia en el local «" & sLocalAlternativo & "» puede retirar toda la mercadería." & vbCrLf & "¿Desea retirar igualmente aquí sólo los posibles?", 9500, Pregunto
'                ElseIf vsGridArt.Rows > 1 Then
'                    loc_ShowMsg "Sugerencia en el local «" & sLocalAlternativo & "» puede retirar toda la mercadería." & vbCrLf & "¿Desea retirar igualmente aquí sólo los posibles?", 9500, Pregunto
'                Else
'                    loc_ShowMsg "Sugerencia en el local «" & sLocalAlternativo & "» también puede retirar toda la mercadería." & vbCrLf & "¿Desea retirar igualmente aquí?", 9500, Pregunto
'                End If
'            Else
'                loc_ShowMsg "¿Desea retirar el/los artículo/s?", 9000, Pregunto
'            End If
'        End If

        loc_ShowMsg "Documento asignado a cola de entrega", 2500, Pregunto
        
    ElseIf vsGrid2.Rows > 0 Then
    
        If bLocalAlternativo Then
            loc_ShowMsg "Atención en el local «" & sLocalAlternativo & "» puede retirar toda la mercadería.", 7000, Advierto
        ElseIf vsGrid2.Tag <> "" Then
            loc_ShowMsg "Verifique los locales donde puede retirar la mercadería.", 9000, Informo
        Else
            loc_ShowMsg "No hay artículos para retirar en el documento.", 9000, Informo
        End If
    Else
        loc_ShowMsg "No hay artículos para retirar para el documento", 6000, Advierto
    End If
    
    imgRetirar.Visible = (vsGridArt.Rows > 0)
'    imgNo.Visible = (vsGridArt.Rows > 0)
'    imgCancelar.Visible = True
    If vsGridArt.Rows = 1 And vsGrid2.Rows = 0 Then
        'Tengo sólo un artículo para entregar --> lo doy x entregado.
        imgRetirar_Click
    Else
        If vsGridArt.Rows > 0 Or vsGrid2.Rows > 0 And paWavLista <> "" Then
            loc_SetSonido paWavLista
        End If
    End If
    Foco tcBarra
    
    If Not vsGridArt.Visible Then vsGrid2.Top = vsGridArt.Top Else vsGrid2.Top = linfoGrid.Top + linfoGrid.Height + 240
    Exit Sub
    
errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los artículos del documento.", Err.Description
End Sub

Private Function fnc_Qstock(ByVal idArt As Long) As Long
Dim mSQL As String
Dim rsA As rdoResultset
On Error GoTo errStock
    
    mSQL = " Select STLCantidad From StockLocal Where STLArticulo = " & idArt & " And STLEstado =" & paEstadoArticuloEntrega & _
           " And STLLocal =" & paCodigoDeSucursal
    Set rsA = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    If Not rsA.EOF Then fnc_Qstock = rsA("STLCantidad")
    rsA.Close
    Exit Function

errStock:
    fnc_Qstock = 0
    'clsGeneral.OcurrioError "Ocurrió un error al verificar la cantidad de stock.", Err.Description
End Function

Private Function fnc_GrabarRetiro() As Long
On Error GoTo errBT
Dim sXMLArticulos As String, sRenglon As String, sRen As String
Dim iQ As Integer

    fnc_GrabarRetiro = -1
    Err.Clear
    rdoErrors.Clear
    
    '(ArtID int, Cant smallint, Est tinyint, SArr smallint)
    sRenglon = "<Ren ArtID=""[mIDArt]"" Cant=""[mQ]"" Est=""[mEst]"" SArr=""[mSArr]""></Ren>"
    With vsGridArt
        For iQ = 0 To .Rows - 1
            'Si tengo Q > 0 estoy entregando.
            If .Cell(flexcpText, iQ, 1) > 0 Then
                sRen = Replace(sRenglon, "[mIDArt]", .Cell(flexcpData, iQ, 0))
                sRen = Replace(sRen, "[mQ]", .Cell(flexcpText, iQ, 1))
                If paArrimar = 1 Then
                    sRen = Replace(sRen, "[mEst]", .Cell(flexcpData, iQ, 3))
                    If .Cell(flexcpData, iQ, 3) = 1 Then
                        sRen = Replace(sRen, "[mSArr]", .Cell(flexcpText, iQ, 1))
                    Else
                        sRen = Replace(sRen, "[mSArr]", 0)
                    End If
                Else
                    sRen = Replace(sRen, "[mSArr]", 0)
                    sRen = Replace(sRen, "[mEst]", 2)
                End If
                sXMLArticulos = sXMLArticulos & sRen
            End If
        Next
    End With
    sXMLArticulos = "<ROOT>" & sXMLArticulos & "</ROOT>"

    Cons = "prg_EntregaCliente_GraboDocumento " & iTipoDoc & ", " & iDocumento & _
            ", '" & Format(gFechaDocumento, "yyyy/mm/dd hh:nn:ss") & "', " & paCodigoDeSucursal & ", '" & sXMLArticulos & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux(0) >= 1 Then
        fnc_GrabarRetiro = RsAux(0)
    Else
        loc_ShowMsg RsAux(1), 2000, Error
    End If
    RsAux.Close
    Exit Function
errBT:
    clsGeneral.OcurrioError "Error", Err.Description, "Grabar"
End Function

Private Sub vsGridArt_KeyDown(KeyCode As Integer, Shift As Integer)

    If Shift <> 0 Or vsGridArt.Rows = 0 Then Exit Sub
    Select Case KeyCode
        Case vbKeyAdd
            If CInt(vsGridArt.Cell(flexcpText, vsGridArt.Row, 1)) < vsGridArt.Cell(flexcpData, vsGridArt.Row, 1) Then
                vsGridArt.Cell(flexcpText, vsGridArt.Row, 1) = CInt(vsGridArt.Cell(flexcpText, vsGridArt.Row, 1)) + 1
             End If
        Case vbKeySubtract
            If CInt(vsGridArt.Cell(flexcpText, vsGridArt.Row, 1)) > 0 Then
                vsGridArt.Cell(flexcpText, vsGridArt.Row, 1) = CInt(vsGridArt.Cell(flexcpText, vsGridArt.Row, 1)) - 1
            End If
    End Select
    
End Sub

Public Function fnc_StockLocalAlternativo(ByVal iIDArt As Long, ByVal iCantidad As Integer) As Boolean
On Error GoTo errVerSL
Dim sQy As String
Dim RsAux As rdoResultset
    
    sQy = "Select STLCantidad " & _
            " From StockLocal " & _
            " Where STlLocal = " & iLocalAlternativo & " And STlCantidad > 0 " & _
            " And STlEstado = " & paEstadoArticuloEntrega & " And STlArticulo = " & iIDArt
    Set RsAux = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If iCantidad < RsAux(0) Then fnc_StockLocalAlternativo = True
    End If
    RsAux.Close
Exit Function
errVerSL:
End Function


Public Function fnc_StockEmergencia(ByVal idArt As Long) As Integer
On Error GoTo errSockEme
    
    bCercano = False
    Cons = "Select * From ArticuloStockEntrega Where ASEArticulo = " & idArt & " And ASELocal = " & paCodigoDeSucursal
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux("ASeCercano")) Then If RsAux("ASeCercano") Then bCercano = True
        If Now < CDate(Date & " " & Format(iHoraAlternativa, "00:00") & ":01") Then
            fnc_StockEmergencia = RsAux("ASEStkEmergencia")
        End If
    End If
    RsAux.Close
    Exit Function
    
errSockEme:
    'clsGeneral.OcurrioError "Ocurrió un error al verificar el stock de emergencia", Err.Description
End Function

Private Sub wbNav_GotFocus()
    tcBarra.SetFocus
    Debug.Print "PASO"
End Sub



