VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMostrador 
   BackColor       =   &H00F5FFFF&
   Caption         =   "Entrega de Mercadería"
   ClientHeight    =   9450
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14730
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMostrador.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9450
   ScaleWidth      =   14730
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList img16 
      Left            =   240
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMostrador.frx":0ECA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer TArr 
      Left            =   240
      Top             =   2640
   End
   Begin VB.TextBox tcBarra 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3720
      MaxLength       =   30
      TabIndex        =   0
      Top             =   7440
      Width           =   3495
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   240
      Top             =   5040
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
            Picture         =   "frmMostrador.frx":1199
            Key             =   "exc"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMostrador.frx":18ED3
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMostrador.frx":30C0D
            Key             =   "inf"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMostrador.frx":48947
            Key             =   "car"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMostrador.frx":60681
            Key             =   "sto"
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsGridArt 
      Height          =   1695
      Left            =   60
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   2990
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483641
      BackColorFixed  =   -2147483630
      ForeColorFixed  =   -2147483639
      BackColorSel    =   -2147483624
      ForeColorSel    =   8421504
      BackColorBkg    =   16777215
      BackColorAlternate=   16119285
      GridColor       =   -2147483635
      GridColorFixed  =   -2147483635
      TreeColor       =   -2147483632
      FloodColor      =   0
      SheetBorder     =   16777215
      FocusRect       =   0
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
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
   Begin VSFlex6DAOCtl.vsFlexGrid vsGrid 
      Height          =   1575
      Left            =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2400
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   2778
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483641
      BackColorFixed  =   -2147483630
      ForeColorFixed  =   -2147483639
      BackColorSel    =   16448
      ForeColorSel    =   6710886
      BackColorBkg    =   -2147483624
      BackColorAlternate=   16119285
      GridColor       =   -2147483635
      GridColorFixed  =   -2147483635
      TreeColor       =   -2147483632
      FloodColor      =   0
      SheetBorder     =   16777215
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   3
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
   Begin VB.Image imgSound 
      Height          =   240
      Left            =   120
      Picture         =   "frmMostrador.frx":60F50
      Top             =   7440
      Width           =   240
   End
   Begin VB.Label lserie 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NroSerie"
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   1080
      TabIndex        =   5
      Top             =   3720
      Width           =   7815
   End
   Begin VB.Shape shSerie 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   1080
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   7935
   End
   Begin VB.Label lingreso 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese usuario o factura:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   7440
      Width           =   3255
   End
   Begin VB.Label lbUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario no logeado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      Top             =   7440
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   7440
      Picture         =   "frmMostrador.frx":6136B
      Top             =   7440
      Width           =   360
   End
   Begin VB.Label lMsg 
      BackColor       =   &H00004080&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   1680
      TabIndex        =   2
      Top             =   8160
      Width           =   11415
   End
   Begin VB.Image ImgIcon 
      Height          =   615
      Left            =   360
      Top             =   8280
      Width           =   735
   End
   Begin VB.Shape ShMsg 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   8040
      Width           =   14415
   End
   Begin VB.Menu MnuGrilla 
      Caption         =   "MnuGrilla"
      Visible         =   0   'False
      Begin VB.Menu MnuAnular 
         Caption         =   "Anular"
      End
      Begin VB.Menu MnuClienteSeFue 
         Caption         =   "ClienteSeFue"
      End
      Begin VB.Menu MnuEnOtroLocal 
         Caption         =   "Cliente retira en otro local"
      End
   End
   Begin VB.Menu MnuSonidos 
      Caption         =   "Sonidos"
      Visible         =   0   'False
      Begin VB.Menu MnuSoundOK 
         Caption         =   "OK"
      End
      Begin VB.Menu MnuSoundMal 
         Caption         =   "Mal"
      End
   End
End
Attribute VB_Name = "frmMostrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'10/3/2008 no hago más borrar productosvendidos lo deje sólo cuando es devolución.
'9/7/2008 se marca con una señal si presiona f7 que el cliente fue llamado

Private sDocsOcultos As String
Private sCliFueLlamado As String

Private iRowSerie As Integer        'Si tengo además de la s en la celda 2 --> tiene que darme ese nro. de serie.
Private iTipoDoc As Integer
Private iDocumento As Long
Private iRemito As Long
Private gFechaDocumento As Date

Private iIDUsuario As Long          'guardo el id del usuario que entrega el artículo

Private iPaso As Integer             ' 1: Lista de facturas 2:Entrega de art de factura
Private lIdArticulo As Long

Private lCodServicio As Long

Private lToleranciaEntrega As Long

Private sWav As String

Private bSinEnAuxiliar As Boolean

Private arrSucesoArt() As Long
Dim oHub As New ClientHub
Private prmHUBMetod As String, prmHUBURL As String, prmHUBNombre As String

Private Enum EstadoR 'estado de entrega
    SinEntregar = 1
    Arrimado = 2
    Entregado = 3
    Anulado = 4
    ClienteSeFue
    ClienteRetiraEnOtroLocal
End Enum

Private Enum eEstMsg
    informo = 0
    Advierto = 1
    Pregunto = 2
    Error = 3
End Enum

Private Type typNroSerie
    Articulo As Long
    NroSerie As String
End Type

Public Enum TipoLocal
    Camion = 1
    Deposito = 2
End Enum

Dim arrArtFecha() As typNroSerie

Dim arrNroSerie() As typNroSerie

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Function pf_FueArrimado(ByVal iTipoDoc As Byte, ByVal iDocID As Long) As Boolean
On Error Resume Next
    pf_FueArrimado = (InStr(1, sDocsOcultos, iTipoDoc & "|" & iDocID & ";"))
End Function

'Private Function pf_ClienteLlamado(ByVal iTipoDoc As Byte, ByVal iDocID As Long) As Boolean
'On Error Resume Next
'    pf_ClienteLlamado = (InStr(1, sCliFueLlamado, iTipoDoc & "|" & iDocID & ";"))
'End Function

Private Function EliminarAccionClienteLlamado()
On Error Resume Next
Dim iRow As Integer, iLast As Integer
    iLast = -1
    With vsGridArt
        For iRow = 0 To vsGridArt.Rows - 1
            If Not IsEmpty(.Cell(flexcpPicture, iRow, 0)) Then 'pf_ClienteLlamado(Val(.Cell(flexcpData, iRow, 0)), Val(.Cell(flexcpData, iRow, 1))) Then
                iLast = iRow
            Else
                Exit For
            End If
        Next
        If iLast >= 0 Then
            'sCliFueLlamado = Replace(sCliFueLlamado, Val(.Cell(flexcpData, iLast, 0)) & "|" & Val(vsGridArt.Cell(flexcpData, iLast, 1)) & ";", "")
            Cons = "UPDATE EntregaAuxiliar SET EAuTiempoTotal = 0 WHERE EAuTipo = " & Val(.Cell(flexcpData, iLast, 0)) _
                & "AND EAuDocumento = " & Val(.Cell(flexcpData, iLast, 1)) & " AND EAuTiempoTotal = -1 AND EAuEstado = 2"
            cBase.Execute Cons
            
            On Error Resume Next
            oHub.InvokeMethod prmHUBMetod
        End If
    End With
    
End Function

Private Sub CargoPrmsSignalR()
On Error GoTo errCPS
Dim sQy As String
Dim rsP As rdoResultset

    sQy = "select ParNombre, ParTexto From Parametro where ParNombre like 'signalr%'"
    Set rsP = cBase.OpenResultset(sQy, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not rsP.EOF
        Select Case LCase(Trim(rsP("ParNombre")))
            Case LCase("signalREntregaTVMetodo")
                prmHUBMetod = Trim(rsP("ParTexto"))
            Case LCase("signalRURL")
                prmHUBURL = Trim(rsP("ParTexto"))
            Case LCase("signalRHubEntregaTV")
                prmHUBNombre = Trim(rsP("ParTexto"))
        End Select
        rsP.MoveNext
    Loop
    rsP.Close
    Exit Sub
errCPS:
clsGeneral.OcurrioError "Error al cargar los parámetros.", Err.Description, "Cargo Parámetros"
End Sub


Private Sub AsignarClienteLlamado()
On Error Resume Next
Dim iRow As Integer, iLast As Integer
    iLast = -1
    With vsGridArt
    'Se puede preguntar tb si el picture = nothing
        For iRow = 0 To vsGridArt.Rows - 1
            If Not IsEmpty(.Cell(flexcpPicture, iRow, 0)) Then   'pf_ClienteLlamado(Val(.Cell(flexcpData, iRow, 0)), Val(.Cell(flexcpData, iRow, 1))) Then
                iLast = iRow
            Else
                Exit For
            End If
        Next
        If iLast = -1 And vsGridArt.Rows > 0 Then
            iLast = 0
        ElseIf iLast >= 0 Then
            iLast = iLast + 1
        End If

        If iLast > -1 Then
'            sCliFueLlamado = sCliFueLlamado & Val(.Cell(flexcpData, iLast, 0)) & "|" & Val(vsGridArt.Cell(flexcpData, iLast, 1)) & ";"
            Cons = "UPDATE EntregaAuxiliar SET EAuTiempoTotal = -1 WHERE EAuTipo = " & Val(.Cell(flexcpData, iLast, 0)) _
                & " AND EAuDocumento = " & Val(.Cell(flexcpData, iLast, 1)) & " AND EAuEstado = 2"
            cBase.Execute Cons
            
            On Error Resume Next
            oHub.InvokeMethod prmHUBMetod
            
        End If
    End With
End Sub

Private Sub ps_ShowNormal()
On Error Resume Next
    tcBarra.Visible = True
    'vsGridArt.ColHidden(0) = False
    vsGridArt.ColWidth(0) = 8000
    vsGrid.Cell(flexcpFontSize, 0, 0, vsGridArt.Rows - 1) = 12
    Form_Resize
    Me.BackColor = &HF5FFFF
    tcBarra.SetFocus
End Sub

Private Sub ps_OcultoDocs(ByVal iTipoDoc As Byte, ByVal iIDDoc As Long)
On Error Resume Next
    'Tomo el id del doc y a partir de el inserto en string
    If iIDDoc = 0 Then Exit Sub
    TArr.Enabled = False
    Dim iQ As Integer
    For iQ = 0 To vsGridArt.Rows - 1
        sDocsOcultos = sDocsOcultos & Val(vsGridArt.Cell(flexcpData, iQ, 0)) & "|" & Val(vsGridArt.Cell(flexcpData, iQ, 1)) & ";"
        If Val(vsGridArt.Cell(flexcpData, iQ, 0)) = iTipoDoc And Val(vsGridArt.Cell(flexcpData, iQ, 1)) = iIDDoc Then
            Exit For
        End If
    Next
    CargoFacturas
    TArr.Enabled = True
End Sub

Private Sub loc_ShowArts(ByVal bVisible As Boolean)
    vsGrid.Rows = 0
    vsGrid.Visible = bVisible
End Sub


Private Sub loc_InitGrid()
    With vsGridArt
        .BackColorBkg = &HFFFFFF
        .ColWidth(0) = 8000
        .Rows = 0
        .FontSize = IIf(paArrimar = 1, 20, 15)
        .RowHeightMin = IIf(paArrimar = 1, 600, 455)
    End With
    With vsGrid
        .BackColorBkg = &HFFFFFF
        .ColWidth(0) = 1200
        .ColWidth(1) = 200
        .Visible = (paArrimar <> 1)
        .Rows = 0
        .RowHeightMin = 455
    End With
End Sub

Private Sub loc_VerificoArrime()
On Error GoTo errVA
Dim iAntes As Byte
    iAntes = paArrimar
    Cons = "Select * from Parametro Where ParNombre IN('dep_Estado_Arrimar_" & paCodigoDeSucursal & "')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then paArrimar = RsAux("ParValor")
    RsAux.Close
    If iAntes <> paArrimar Then loc_InitGrid: Form_Resize
    Exit Sub
errVA:
    clsGeneral.OcurrioError "Error al cargar el parámetro del arrime.", Err.Description, "Verifico arrime"
End Sub

Private Function fnc_GetColorOrden(ByVal iOrden As Integer) As String
    
    Select Case (iOrden + 1) Mod 11
        Case 1: fnc_GetColorOrden = &H99CC01
        Case 2: fnc_GetColorOrden = &HC4E3F7
        Case 3: fnc_GetColorOrden = &H1CCFF
        Case 4: fnc_GetColorOrden = &HCC9933
        Case 5: fnc_GetColorOrden = &H839F2F
        Case 6: fnc_GetColorOrden = &HAFBDF6
        Case 7: fnc_GetColorOrden = &HF7E3C4
        
        Case 8: fnc_GetColorOrden = &H3399FF
        Case 9: fnc_GetColorOrden = &HCACACA
        Case 10: fnc_GetColorOrden = &HFF9966   '&H986601
        
        Case 11: fnc_GetColorOrden = &H99CCFF
    End Select
    
End Function


Private Sub loc_Cancel()
    lingreso.Caption = "Ingrese su usuario:"
    lbUsuario.Caption = "Usuario no logeado"
    iIDUsuario = 0
    ReDim arrSucesoArt(0)
    CargoFacturas
End Sub

Private Function fnc_EsParcial() As Boolean
Dim iQ As Integer
    fnc_EsParcial = False
    With vsGridArt
        For iQ = .FixedRows To .Rows - 1
            If Val(.Cell(flexcpText, iQ, 1)) <> Val(.Cell(flexcpData, iQ, 1)) Then fnc_EsParcial = True: Exit For
        Next
    End With
End Function

Private Function fnc_EsDevolucion() As Boolean
    fnc_EsDevolucion = (iTipoDoc = TipoDocumento.NotaCredito Or iTipoDoc = TipoDocumento.NotaDevolucion Or iTipoDoc = TipoDocumento.NotaEspecial)
End Function

Private Sub loc_ConsultoGrabarParcial()

    If iPaso = 2 Or iPaso = 3 Then
        If fnc_EsParcial Then
            If fnc_EsDevolucion Then
                If MsgBox("Esto es una recepción de mercadería por devolución." & Chr(vbKeyReturn) & "El cliente debe devolver todos los artículos, de lo contrario no podrá realizar el ingreso." & vbCrLf & vbCrLf & "¿Desea cancelar la devolución?", vbExclamation + vbYesNo, "Faltan Artículos") = vbYes Then
                    loc_Cancel
                End If
            Else
                If MsgBox("¿Desea confirmar la Entrega Parcial?", vbQuestion + vbYesNo, "ENTREGA PARCIAL") = vbYes Then
                    loc_AccionFinalizar True
                Else
                    If MsgBox("Ud. no le va a entregar ningún producto al cliente." & vbCrLf & vbCrLf & "¿Confirma cancelar toda la entrega?", vbQuestion + vbYesNo) = vbYes Then
                        loc_Cancel
                    End If
                End If
            End If
        Else
            loc_Cancel
        End If
    Else
        loc_Cancel
    End If
End Sub

Private Sub loc_Help()
On Error GoTo errHelp
    Screen.MousePointer = 11
    Dim aFile As String
    Cons = "Select * from Aplicacion Where AplNombre = '" & App.Title & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux!AplHelp) Then aFile = Trim(RsAux!AplHelp)
    RsAux.Close
    If aFile <> "" Then EjecutarApp aFile
    Screen.MousePointer = 0
    Exit Sub
errHelp:
    clsGeneral.OcurrioError "Error al activar el archivo de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function fnc_BuscoDocPorTexto(adTexto As String, retIDDoc As Long, retIDTipoD As Byte) As Boolean
On Error GoTo errDoc
        
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
               " And DocTipo IN (" & TipoDocumento.Contado & ", " & TipoDocumento.Credito & ", " & TipoDocumento.Remito & ", " & TipoDocumento.NotaCredito & ", " & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial & ")" & _
               " And DocAnulado = 0"
        
    'Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If ObtenerResultSet(cBase, RsAux, Cons, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Function
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
        fnc_BuscoDocPorTexto = True
        retIDDoc = adCodigo
        retIDTipoD = adTipoD
    Else
        loc_ShowMsg "No se encontró un documento de compra.", 3000, Advierto
    End If
    Screen.MousePointer = 0
    Exit Function
    
errDoc:
    loc_ShowMsg "Error al buscar el documento: " & Err.Description, 3000, Error
    Screen.MousePointer = 0
End Function

Private Sub loc_ShowMsg(ByVal sTxt As String, ByVal iInterval As Integer, ByVal iEstado As eEstMsg)
On Error Resume Next
Dim iFColor As Long

    Select Case iEstado
        Case 0  'información
            ImgIcon.Picture = imlIcons.ListImages("inf").Picture
            ShMsg.FillColor = &HCEBAB3
            ShMsg.BorderColor = &HFAF0EB
            loc_SetSonido "entregaok.wav"
        
        Case 1  'advertencia
            ImgIcon.Picture = imlIcons.ListImages("exc").Picture
            ShMsg.FillColor = &HD0FFFF
            ShMsg.BorderColor = &HC0C0&
            lMsg.ForeColor = vbWhite
            loc_SetSonido "entregamal.wav"
            
        Case 2  'Pregunta
            ImgIcon.Picture = imlIcons.ListImages("pre").Picture
            ShMsg.FillColor = &HCEBAB3
            ShMsg.BorderColor = &HFAF0EB
            
        Case 3  'Error
            ImgIcon.Picture = imlIcons.ListImages("sto").Picture
            ShMsg.FillColor = &HFFFFFF
            ShMsg.BorderColor = &HC0&
            lMsg.ForeColor = vbWhite
    End Select
    lMsg.Caption = sTxt
    
    lMsg.Visible = True
    ImgIcon.Visible = True
    ShMsg.Visible = True
    
    TArr.Interval = iInterval
    TArr.Enabled = True
End Sub

Private Sub loc_AccionFinalizar(Optional ByVal bParcial As Boolean = False)
Dim bVisible As Boolean
    
    bVisible = fnc_HayFilasVisibles
    
    
    If Not bVisible Or bParcial Then
         Dim objSuceso As New clsSuceso
        'Verifico si tengo que grabar suceso
        If Not fnc_EsDevolucion Then
            If UBound(arrSucesoArt) > 0 Then
                Dim aUsuario As Long
                Dim sDefensa As String
                'Llamo al registro del Suceso-------------------------------------------------------------
                Set objSuceso = New clsSuceso
                Do
                    aUsuario = 0
                    loc_SetSonido "entregamal.wav"
                    objSuceso.TipoSuceso = 24
                    objSuceso.ActivoFormulario iIDUsuario, "No paso código de barras del artículo", cBase
                    Me.Refresh
                    aUsuario = objSuceso.Usuario
                    sDefensa = objSuceso.Defensa
                    Set objSuceso = Nothing
                    If aUsuario = 0 Then
                        Screen.MousePointer = 0
                        If MsgBox("NO PODRÁ GRABAR LA ENTREGA si no escribe el suceso." & vbCrLf & vbCrLf & "¿Desea CANCELAR la entrega?", vbQuestion + vbYesNo + vbDefaultButton2, "Atención") = vbYes Then
                            loc_Cancel
                            Exit Sub
                        End If
                    End If
                Loop Until aUsuario > 0
                
            End If
        End If
        
        
        'Verifico si tengo que pedir suceso por entregar un artículo fuera de fecha
        Dim sArtFR As String
        Dim iQ As Integer
        For iQ = 1 To UBound(arrArtFecha)
            If arrArtFecha(iQ).Articulo > 0 And arrArtFecha(iQ).NroSerie <> "" Then
                'Recorro la grilla y me fijo si al final lo está entregando.
                With vsGridArt
                    For I = 0 To .Rows - 1
                        If (.RowHidden(I) Or (Val(.Cell(flexcpData, I, 1)) <> Val(.Cell(flexcpText, I, 1)))) And Val(.Cell(flexcpData, I, 0)) = arrArtFecha(iQ).Articulo Then
                            sArtFR = sArtFR & IIf(sArtFR <> "", ", ", "") & arrArtFecha(iQ).Articulo
                            Exit For
                        End If
                    Next
                End With
            End If
        Next iQ
        
        If sArtFR <> "" Then
            Dim usuariosuceso As Integer
            Dim strdefensa As String
            
            'Llamo al registro del Suceso-------------------------------------------------------------
            Set objSuceso = New clsSuceso
            Do
                usuariosuceso = 0
                loc_SetSonido "entregamal.wav"
                objSuceso.TipoSuceso = 24
                objSuceso.ActivoFormulario iIDUsuario, "Artículos entregados fuera del rango permitido", cBase
                Me.Refresh
                usuariosuceso = objSuceso.Usuario
                strdefensa = objSuceso.Defensa
                Set objSuceso = Nothing
                If usuariosuceso = 0 Then
                    Screen.MousePointer = 0
                    If MsgBox("NO PODRÁ GRABAR LA ENTREGA si no escribe el suceso." & vbCrLf & vbCrLf & "¿Desea CANCELAR la entrega?", vbQuestion + vbYesNo + vbDefaultButton2, "Atención") = vbYes Then
                        loc_Cancel
                        Exit Sub
                    End If
                End If
            Loop Until usuariosuceso > 0
        End If
        
        Dim bRet As Boolean
        bRet = AccionGrabar(aUsuario, sDefensa, usuariosuceso, strdefensa, sArtFR)
        iPaso = 1
        iIDUsuario = 0
        lbUsuario.Caption = " "
        lingreso.Caption = "Ingrese usuario o factura:"
        ReDim arrSucesoArt(0)
        
        If bRet Then
            loc_ShowMsg "Se almacenaron los datos", 700, informo
        Else
            loc_ShowMsg "Visualice en detalle de facturas antes de reintentar.", 6000, Error
        End If
        Me.BackColor = &HF5FFFF
    End If
End Sub

Private Sub loc_LimpiarMsg()
    lMsg.Caption = ""
    ImgIcon.Visible = False
    ShMsg.Visible = False
End Sub

Private Sub loc_SetSonido(ByVal sFile As String)
On Error Resume Next
Dim Result As Long
    Result = sndPlaySound(sWav & sFile, 1)
End Sub

Private Sub loc_FocoTBarra()
On Error Resume Next
    With tcBarra
        If .Enabled Then
            .SelStart = 0: .SelLength = Len(.Text): .SetFocus
        End If
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyEscape Then
        '23/7/2007 si ya tengo ingresado algún artículo le consulto si desea cancelar o no
        loc_ConsultoGrabarParcial
        Me.BackColor = &HF5FFFF
        If iPaso = 1 And Not tcBarra.Visible Then ps_ShowNormal
    ElseIf KeyCode = vbKeyF5 And paArrimar = 0 Then
        'Veo en que paso está
        If iPaso = 1 Then
            'Reubico las grillas.
            On Error Resume Next
            If tcBarra.Visible Then
                'vsGridArt.ColHidden(0) = True
                vsGridArt.ColWidth(0) = 5000
                vsGridArt.FontSize = 14
                vsGridArt.Move 0, 0, 7000, Me.ScaleHeight
                'vsGridArt.Font.Size = 12
                vsGrid.Move 7005, 0, Me.ScaleWidth - 7005, Me.ScaleHeight
                tcBarra.Visible = False
                CargoFacturas
                vsGrid.SetFocus
            Else
                sDocsOcultos = ""
                ps_ShowNormal
            End If
        Else
            MsgBox "Presione <ESC> para cancelar la entrega pendiente.", vbInformation, "Atención"
        End If
    ElseIf KeyCode = vbKeyF7 And iPaso = 1 Then
        '9/7/2008 agregado.
        'Le pongo al primer renglón un icono que indica que el cliente ya fue llamado.
        EliminarAccionClienteLlamado
        CargoFacturas
    ElseIf KeyCode = vbKeyF8 And iPaso = 1 Then
        AsignarClienteLlamado
        CargoFacturas
    ElseIf KeyCode = vbKeyF9 And iPaso = 1 Then
        'MostrarInformacionACliente
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    ObtengoSeteoForm Me, 100, 100
    If Me.Left < 0 Then Me.Left = 100
    If Me.Top < 0 Then Me.Top = 100
    
    loc_InitGrid
    
    ShMsg.Visible = False
    shSerie.Visible = False
    lserie.Visible = False
    iPaso = 1
    
    ReDim arrSucesoArt(0)
    sDocsOcultos = ""
    
    ChDir App.Path
    ChDir ("..")
    If Dir(CurDir & "\Sonidos", vbDirectory) <> "" Then
        sWav = CurDir & "\Sonidos\"
    Else
        MsgBox "No encontré la carpeta SONIDOS en " & CurDir, vbExclamation, "IMPORTANTE"
    End If
    
    Cons = "SELECT IsNull(LocToleranciaEntrega, 0) FROM Local WHERE LocCodigo = " & paCodigoDeSucursal
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        lToleranciaEntrega = RsAux(0)
    End If
    RsAux.Close
        
    CargoFacturas
    
    ConectoSignalr
    
End Sub

Private Sub ConectoSignalr()
On Error GoTo errC
    CargoPrmsSignalR
    Set oHub = New ClientHub
    If Not oHub.ConnectHub(prmHUBURL, prmHUBNombre) Then
        MsgBox "No conectó el signalR"
    End If
    Exit Sub
errC:
    clsGeneral.OcurrioError "Error al conectar.", Err.Description, "SignalR"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    ShMsg.Move ShMsg.Left, Me.ScaleHeight - ShMsg.Height - 360, Me.ScaleWidth - (ShMsg.Left * 2)
    ImgIcon.Move ImgIcon.Left, ShMsg.Top + 240
    lMsg.Move ImgIcon.Left + 1320, ShMsg.Top + 120, Me.ScaleWidth - (ImgIcon.Left + 1560)
    
    tcBarra.Top = ShMsg.Top - 600
    Image1.Top = tcBarra.Top
    lbUsuario.Top = tcBarra.Top
    lingreso.Top = tcBarra.Top
    
    If iPaso = 1 Then tcBarra.Visible = True
    
    If paArrimar <> 1 Then
        vsGridArt.Move vsGridArt.Left, vsGridArt.Top, (Me.ScaleWidth - vsGridArt.Left * 2), ((tcBarra.Top - vsGridArt.Top - 120) / 2)
        vsGrid.Move vsGridArt.Left, vsGridArt.Top + vsGridArt.Height + 60, (Me.ScaleWidth - vsGridArt.Left * 2), tcBarra.Top - (vsGridArt.Top + vsGridArt.Height + 120)
    Else
        vsGridArt.Move vsGridArt.Left, vsGridArt.Top, (Me.ScaleWidth - vsGridArt.Left * 2), tcBarra.Top - vsGridArt.Top - 120
    End If
    vsGridArt.FontSize = IIf(paArrimar = 1, 20, 15)
    vsGridArt.ColWidth(0) = 8000
    
    shSerie.Move ((Me.ScaleWidth - shSerie.Width) / 2), 3500
    lserie.Move shSerie.Left, shSerie.Top + 360
    
    imgSound.Top = Image1.Top
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    GuardoSeteoForm Me
    Set clsGeneral = Nothing
    cBase.Close
    eBase.Close
End Sub



Private Sub imgSound_Click()
    PopupMenu MnuSonidos
End Sub

Private Sub MnuAnular_Click()
    CambiarEstadoArt EstadoR.Anulado, vsGridArt.Cell(flexcpData, vsGridArt.Row, 1), vsGridArt.Cell(flexcpData, vsGridArt.Row, 0)
End Sub

Private Sub MnuClienteSeFue_Click()
    CambiarEstadoArt EstadoR.ClienteSeFue, vsGridArt.Cell(flexcpData, vsGridArt.Row, 1), vsGridArt.Cell(flexcpData, vsGridArt.Row, 0)
End Sub

Private Sub MnuEnOtroLocal_Click()
    CambiarEstadoArt EstadoR.ClienteRetiraEnOtroLocal, vsGridArt.Cell(flexcpData, vsGridArt.Row, 1), vsGridArt.Cell(flexcpData, vsGridArt.Row, 0)
End Sub

Private Sub MnuSoundMal_Click()
    MsgBox "Archivo: " & sWav & "EntregaMal.wav"
    loc_SetSonido "entregamal.wav"
End Sub

Private Sub MnuSoundOK_Click()
    MsgBox "Archivo: " & sWav & "EntregaOK.wav"
    loc_SetSonido "entregaOK.wav"
End Sub

Private Sub TArr_Timer()
On Error Resume Next
    TArr.Enabled = False
    If TArr.Tag = "" Then
        TArr.Tag = "1"
    Else
        TArr.Tag = ""
        loc_VerificoArrime
    End If
    loc_LimpiarMsg
    If iPaso = 1 Then CargoFacturas True
    If LCase(Me.ActiveControl.Name) <> "tcbarra" Then tcBarra.SetFocus
End Sub

Private Function fnc_HayFilasVisibles() As Boolean
Dim iQ As Integer
    fnc_HayFilasVisibles = False
    For iQ = vsGridArt.FixedRows To vsGridArt.Rows - 1
        If Not vsGridArt.RowHidden(iQ) Then fnc_HayFilasVisibles = True: Exit For
    Next
End Function

Private Sub tcBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then loc_Help
End Sub

Private Sub tcBarra_KeyPress(KeyAscii As Integer)
On Error Resume Next
    
    If KeyAscii = vbKeyReturn And Trim(tcBarra.Text) <> "" Then
        loc_LimpiarMsg
        TArr.Enabled = False
        Select Case iPaso
            Case 1
                Me.BackColor = &HF5FFFF
                FormatoBarras tcBarra.Text
                
            Case 2
                If iIDUsuario = 0 Then
                    If ControloUsuario(Mid(tcBarra.Text, 2)) Then
                        iPaso = 2: lingreso.Caption = "Artículo:"
                        'le cambio el color al form.
                        Me.BackColor = &HC0FFFF
                    End If
                Else
                    Me.BackColor = &HC0FFFF
                    BuscarArt tcBarra.Text
                End If
            
            Case 3
                If arrAgregoElemento(lIdArticulo, tcBarra.Text) Then
                    shSerie.Visible = False
                    lserie.Visible = False
                    vsGridArt.Visible = True
                    lingreso.Caption = "Artículo:"
                    iPaso = 2
                    loc_AccionFinalizar
                    tcBarra.SetFocus
                End If
                
            Case 4
                If iIDUsuario = 0 Then
                    If ControloUsuario(Mid(tcBarra.Text, 2)) Then iPaso = 4: lingreso.Caption = "Código de servicio"
                Else
                    If Not lCodServicio = Mid(tcBarra.Text, 2) Then
                        loc_ShowMsg "Los codigos de servicios no son iguales, verifique", 3000, Advierto
                        loc_FocoTBarra
                    Else
                        CambiarEstado Mid(tcBarra.Text, 2)
                    End If
                End If
        End Select
        tcBarra.Text = ""
        
    End If
End Sub

Private Sub FormatoBarras(Texto As String)

Dim iAuxTipo As Byte
Dim iCodDoc As Long
Dim iDBarCode As String
    
    On Error GoTo errInt
    
    TArr.Enabled = False
    Texto = UCase(Texto)
    
    If InStr(1, Texto, "D", vbTextCompare) > 1 Then
    
        If Not IsNumeric(Mid(Texto, 1, InStr(Texto, "D") - 1)) Then loc_ShowMsg "El dato ingresado no es un documento de compra.", 5000, Advierto: loc_FocoTBarra: Exit Sub
        If IsNumeric(Mid(Texto, 1, InStr(Texto, "D") - 1)) And IsNumeric(Trim(Mid(Texto, InStr(Texto, "D") + 1, Len(Texto)))) Then
            iAuxTipo = CLng(Mid(Texto, 1, InStr(Texto, "D") - 1))
            iCodDoc = CLng(Trim(Mid(Texto, InStr(Texto, "D") + 1, Len(Texto))))
        End If
        
    ElseIf InStr(1, Texto, "U", vbTextCompare) > 0 Then
        'Controlo id de usuario
        iDBarCode = CStr(Trim(Mid(Texto, InStr(Texto, "U") + 1, Len(Texto))))
        If Not loc_BuscarUsuario(iDBarCode) Then loc_ShowMsg "El usuario ingresado no existe.", 3000, Advierto: loc_FocoTBarra: Exit Sub
        
        lingreso.Caption = "Ingrese una factura:"
        loc_FocoTBarra
        Exit Sub
    
    ElseIf LCase(Left(tcBarra.Text, 1)) = "s" And Len(tcBarra.Text) > 1 Then
        iCodDoc = Mid(tcBarra.Text, 2)
        lCodServicio = Mid(tcBarra.Text, 2)
        loc_BuscoServicio iCodDoc, 1
        Exit Sub
    Else
        'pudo poner serie y número
        If IsNumeric(Texto) Then
            iAuxTipo = TipoDocumento.Remito        'Remito
            iCodDoc = Texto
        Else
            'Puso Serie y Numero de Documento
             If Not fnc_BuscoDocPorTexto(Texto, iCodDoc, iAuxTipo) Then iCodDoc = -1
        End If
        
    End If
   
    Select Case iAuxTipo
        Case TipoDocumento.Remito
             BuscoRemito iCodDoc
        
        Case TipoDocumento.Contado, TipoDocumento.Credito, TipoDocumento.Remito
             BuscoDocumento iAuxTipo, iCodDoc
            
        Case TipoDocumento.NotaCredito, TipoDocumento.NotaDevolucion, TipoDocumento.NotaEspecial
            BuscoDocumento iAuxTipo, iCodDoc
            
        Case Else
            If iIDUsuario <> 0 Then
                loc_ShowMsg "El dato ingresado no es un documento de compra.", 5000, Advierto
                loc_FocoTBarra
            End If
    End Select
    Exit Sub
errInt:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error en el formato de barras.", Err.Description
End Sub
Private Function NoRepetido(ByVal iTipo As Long, iDocumento As Long) As Boolean
On Error Resume Next
Dim I As Integer
    With vsGridArt
        For I = 0 To .Rows - 1
            If .Cell(flexcpData, I, 0) = iTipo And .Cell(flexcpData, I, 1) = iDocumento Then NoRepetido = True: Exit For
        Next
    End With
End Function

Private Sub loc_InsertGridDocumento(ByVal iTipoDoc As Long, ByVal iDoc As Long, ByVal sNomCli As String, ByVal sDoc As String, ByVal bAplicoColor As Boolean, ByVal ClienteFueLlamado As Boolean)
Dim iAux As Long
    With vsGridArt
        .AddItem sNomCli
        
        .Cell(flexcpData, .Rows - 1, 0) = iTipoDoc
        .Cell(flexcpData, .Rows - 1, 1) = iDoc

        .Cell(flexcpText, .Rows - 1, 1) = sDoc
        If bAplicoColor Then .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = fnc_GetColorOrden(.Rows - 1)
        
        If ClienteFueLlamado Then    'pf_ClienteLlamado(iTipoDoc, iDoc) Then
            .Cell(flexcpPicture, .Rows - 1, 0) = img16.ListImages(1).Picture
        End If
        
    End With
End Sub

Private Sub loc_CargoFacturasyArticulos()
Dim Cons As String, sDoc As String
Dim RsAux As rdoResultset
Dim iAux As Long, iCodAnt As Long

    loc_ShowArts True

    Cons = " Select NomCli = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1))+ " _
           & " RTrim(' ' + CPeNombre2), NomEmp = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')'), CliTipo, 0 SerCodigo, DocSerie, Docnumero, EAuFechaHora, EAuDocumento, EAuTipo, " _
           & " ArtId, IsNull(AEsNombre, ArtNombre) ArtNombre, ArtTipo, EAuArticulo, EAuCantidad,  RenARetirar ARetirar, ArtNroSerie, AEsID, AEsNroSerie, CASE EAuTiempoTotal WHEN -1 THEN 1 ELSE 0 END ClienteLlamado" _
           & " From EntregaAuxiliar Inner Join Documento On EntregaAuxiliar.EauDocumento = Documento.DocCodigo  " _
           & " Inner join Renglon on  EntregaAuxiliar.EAuDocumento = Renglon.RenDocumento" _
           & " Inner join Articulo on articulo.ArtId = renglon.RenArticulo " _
           & " Left Outer Join ArticuloEspecifico ON AEsArticulo = EAuArticulo And 1 = AEsTipoDocumento And AEsDocumento = EAuDocumento" _
           & " Left outer join Cliente on documento.DocCliente = Cliente.CliCodigo " _
           & " Left Outer Join Cpersona on Cliente.CliCodigo = Cpersona.CpeCliente  Left Outer Join CEmpresa " _
           & " On CliCodigo = CEmCliente Where EAuEstado = 2 And EAuTipo <> 0 AND EAuLocal = " & paCodigoDeSucursal _
           & " And EAuDocumento Not IN (Select Distinct(EAuDocumento) from EntregaAuxiliar Where EAutipo in (1,2,6) And EAuEstado = 1) AND DocAnulado = 0 " _
           & " And Renglon.RenArticulo = EntregaAuxiliar.EAuArticulo And RenDocumento = DocCodigo "
    
    
    Cons = Cons & " UNION ALL " _
           & " Select NomCli = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1))+ " _
           & " RTrim(' ' + CPeNombre2), NomEmp = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')'), CliTipo, SerCodigo, DocSerie, Docnumero, EAuFechaHora, EAuDocumento, EAuTipo, " _
           & " ArtId, ArtNombre, ArtTipo, EAuArticulo, EAuCantidad,  0 ARetirar, ArtNroSerie, Null AEsID, '', CASE EAuTiempoTotal WHEN -1 THEN 1 ELSE 0 END ClienteLlamado" _
           & " From EntregaAuxiliar Inner Join Documento On EntregaAuxiliar.EauDocumento = Documento.DocCodigo  " _
           & " Inner Join Servicio On SerDocumento = DocCodigo and EAuArticulo = 0 " _
           & " Inner join Producto ON SerProducto = ProCodigo " _
           & " Inner join Articulo on ProArticulo = ArtID " _
           & " Left outer join Cliente on documento.DocCliente = Cliente.CliCodigo " _
           & " Left Outer Join Cpersona on Cliente.CliCodigo = Cpersona.CpeCliente  Left Outer Join CEmpresa " _
           & " On CliCodigo = CEmCliente Where EAuEstado = 2 And EAuTipo <> 0 AND EAuLocal = " & paCodigoDeSucursal _
           & " And EAuDocumento Not IN (Select Distinct(EAuDocumento) from EntregaAuxiliar Where EAutipo in (1,2,6) And EAuEstado = 1) AND DocAnulado = 0 " _
           & " And EntregaAuxiliar.EAuArticulo = 0 "
    
'    Cons = Cons & " UNION ALL " _
'            & "SELECT NomCli = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1))+  RTrim(' ' + CPeNombre2), NomEmp = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')'), CliTipo, 0 SerCodigo, 'REM', RemCodigo, EAuFechaHora, EAuDocumento, EAuTipo, " _
'            & " ArtId,  ArtNombre, ArtTipo, EAuArticulo, EAuCantidad,  RReAEntregar ARetirar, ArtNroSerie, Null AEsID, '' " _
'            & "FROM EntregaAuxiliar Inner Join Remito On EntregaAuxiliar.EauDocumento = Remito.RemCodigo INNER JOIN Documento ON Remito.RemDocumento = Documento.DocCodigo " _
'            & "Inner join RenglonRemito ON  EntregaAuxiliar.EAuDocumento = RenglonRemito.RReRemito " _
'            & "Inner join Articulo ON articulo.ArtId = renglonRemito.RReArticulo " _
'            & "Left Outer join Cliente on documento.DocCliente = Cliente.CliCodigo " _
'            & "LEFT Outer Join Cpersona on Cliente.CliCodigo = Cpersona.CpeCliente Left Outer Join CEmpresa  On CliCodigo = CEmCliente " _
'            & "WHERE EAuEstado = 2 And EAutipo = 6 AND EAuLocal = " & paCodigoDeSucursal _
'            & " AND EAuDocumento Not IN (Select Distinct(EAuDocumento) From EntregaAuxiliar Where EAutipo = 6 And EAuEstado = 1) " _
'            & "AND DocAnulado = 0 And renglonRemito.RReArticulo = EntregaAuxiliar.EAuArticulo"
           
           
'20/10/2016
'    Cons = Cons & " Union All " _
'           & " Select NomCli = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1))+ RTrim(' ' + CPeNombre2) " _
'           & ", NomEmp = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')'), CliTipo, SerCodigo, '' as DocSerie, 0 As DocNumero, EAuFechaHora, EAuDocumento, EAuTipo, " _
'           & " ArtId,  ArtNombre, ArtTipo, EAuArticulo, EAuCantidad,  1 ARetirar, '', Null AEsID, '', CASE EAuTiempoTotal WHEN -1 THEN 1 ELSE 0 END ClienteLlamado" _
'           & " From EntregaAuxiliar, Servicio" _
'           & " Left outer join Cliente on serCliente = Cliente.CliCodigo " _
'           & " Left Outer Join Cpersona on Cliente.CliCodigo = Cpersona.CpeCliente  Left Outer Join CEmpresa " _
'           & " On CliCodigo = CEmCliente " _
'           & ", Producto, Articulo" _
'           & " WHERE EAuEstado = 2 and EAuTipo = 0 AND EAuLocal = " & paCodigoDeSucursal _
'           & " AND EAuDocumento = SerCodigo And SerProducto = ProCodigo And ProArticulo = ArtID" _
'           & " Order by EAuFechaHora "


    Cons = Cons & " Order by EAuFechaHora "

    'Set RsAux = cBase.OpenResultset(Cons) ', rdOpenForwardOnly, rdConcurReadOnly)
    If ObtenerResultSet(cBase, RsAux, Cons, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
    If RsAux.EOF Then sCliFueLlamado = ""
    
    Do While Not RsAux.EOF
    
        If Not pf_FueArrimado(RsAux("EAuTipo"), RsAux("EAuDocumento")) Then
        
        
            If iCodAnt <> RsAux("EAuDocumento") Then
                iCodAnt = RsAux("EAuDocumento")
                If Not NoRepetido(RsAux("EAuTipo"), RsAux("EAuDocumento")) Then
                    If RsAux!CliTipo = 1 Then
                        If Not IsNull(RsAux("NomCli")) Then Cons = Trim(RsAux("NomCli"))
                    Else
                        If Not IsNull(RsAux("NomEmp")) Then Cons = Trim(RsAux("NomEmp"))
                    End If
                    
                    
                    If RsAux("EAuArticulo") = 0 Then
                        If Not IsNull(RsAux("DocNumero")) Then
                            If RsAux("DocNumero") > 0 Then
                                sDoc = "Servicio " & RsAux("SerCodigo") & " fac: " & RsAux("DocSerie") & "-" & RsAux("DocNumero")
                            Else
                                sDoc = "Servicio " & RsAux("SerCodigo")
                            End If
                        Else
                            sDoc = "Servicio " & RsAux("SerCodigo")
                        End If
                    ElseIf Not IsNull(RsAux("DocSerie")) Then
                        sDoc = Trim(RsAux("DocSerie")) & "-" & RsAux("DocNumero")
                    End If
                    loc_InsertGridDocumento RsAux("EAuTipo"), RsAux("EAuDocumento"), Cons, sDoc, True, (RsAux("ClienteLlamado") = 1)
                End If
            End If
            
            'Inserto los artículos
            With vsGrid
                .AddItem IIf(RsAux("Aretirar") > 0, RsAux("ARetirar"), "")
                
                If Not IsNull(RsAux("AEsID")) Then
                    .Cell(flexcpText, .Rows - 1, 2) = "E" & Trim(RsAux("AEsID")) & ":" & Trim(RsAux("ArtNombre"))
                Else
                    .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux("ArtNombre"))
                End If
                .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = vsGridArt.Cell(flexcpBackColor, vsGridArt.Rows - 1, 0)
                
                .Cell(flexcpData, .Rows - 1, 0) = Val(RsAux("EAuTipo"))
                .Cell(flexcpData, .Rows - 1, 1) = Val(RsAux("EAuDocumento"))
        
            End With
        End If
        
        RsAux.MoveNext
    Loop
    RsAux.Close


End Sub

Private Sub CargoFacturas(Optional bEsTimer As Boolean = False)
Dim Cons As String, sDoc As String
Dim RsAux As rdoResultset
Dim iAux As Long
Dim bSonido As Boolean
On Error GoTo errCargar

    lserie.Visible = False
    shSerie.Visible = False
    bSonido = (vsGridArt.Rows = 0 And bEsTimer)
    With vsGridArt
        .Tag = "Docs"
        .ColHidden(2) = True
        .ColHidden(3) = True
        .Rows = 0
        .Visible = True
    End With
    loc_ShowArts True
    vsGridArt.BackColorBkg = &HFFFFFF
    iDocumento = 0
    iRemito = 0
    lbUsuario.Caption = "Usuario no logeado"
    iPaso = 1
    bSinEnAuxiliar = False
    Erase arrNroSerie
    ReDim arrNroSerie(0)
        
    
        
    If paArrimar = 0 Then
        loc_CargoFacturasyArticulos
        GoTo evExit
    End If
    
    
    
    Cons = " Select NomCli = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1))+ " _
           & " RTrim(' ' + CPeNombre2), NomEmp = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')'), CliTipo, SerCodigo, EAuArticulo, DocSerie, Docnumero, EAuFechaHora, EAuDocumento, EAuTipo, CASE EAuTiempoTotal WHEN -1 THEN 1 ELSE 0 END ClienteLlamado " _
           & " From EntregaAuxiliar Inner Join Documento On EntregaAuxiliar.EauDocumento = Documento.DocCodigo  " _
           & " Left Outer Join Servicio On SerDocumento = DocCodigo and EAuArticulo = 0 " _
           & " Left outer join Cliente on documento.DocCliente = Cliente.CliCodigo " _
           & " Left Outer Join Cpersona on Cliente.CliCodigo = Cpersona.CpeCliente  Left Outer Join CEmpresa " _
           & " On CliCodigo = CEmCliente " _
           & " WHERE EAuEstado = 2 And EAuTipo <> 0 AND EAuLocal = " & paCodigoDeSucursal _
           & " AND EAuDocumento Not IN (Select Distinct(EAuDocumento) from EntregaAuxiliar Where EAutipo in (1,2,6) And EAuEstado = 1) AND DocAnulado = 0 "
    
'    Cons = Cons & " UNION ALL " _
'            & "SELECT NomCli = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1))+  RTrim(' ' + CPeNombre2), NomEmp = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')'), CliTipo, 0, EAuArticulo, 'REM', RemCodigo, EAuFechaHora, EAuDocumento, EAuTipo " _
'            & "FROM EntregaAuxiliar Inner Join Remito On EntregaAuxiliar.EauDocumento = Remito.RemCodigo INNER JOIN Documento ON Remito.RemDocumento = Documento.DocCodigo " _
'            & "Left outer join Cliente on documento.DocCliente = Cliente.CliCodigo " _
'            & "LEFT Outer Join Cpersona on Cliente.CliCodigo = Cpersona.CpeCliente Left Outer Join CEmpresa  On CliCodigo = CEmCliente " _
'            & "WHERE EAuEstado = 2 And EAutipo = 6 AND EAuLocal = " & paCodigoDeSucursal _
'            & " AND EAuDocumento Not IN (Select Distinct(EAuDocumento) From EntregaAuxiliar Where EAutipo = 6 And EAuEstado = 1) " _
'            & "AND DocAnulado = 0"
           
    Cons = Cons & " Union All " _
           & " Select NomCli = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1))+ RTrim(' ' + CPeNombre2), " _
           & " NomEmp = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')'), CliTipo, SerCodigo,0 as EAuArticulo, '' as DocSerie, 0 As DocNumero, EAuFechaHora, EAuDocumento, EAuTipo, CASE EAuTiempoTotal WHEN -1 THEN 1 ELSE 0 END ClienteLlamado" _
           & " From EntregaAuxiliar, Servicio" _
           & " Left outer join Cliente on serCliente = Cliente.CliCodigo " _
           & " Left Outer Join Cpersona on Cliente.CliCodigo = Cpersona.CpeCliente  Left Outer Join CEmpresa " _
           & " On CliCodigo = CEmCliente WHERE EAuEstado = 2 and EAuTipo = 0 AND EAuDocumento = SerCodigo AND EAuLocal = " & paCodigoDeSucursal _
           & " Order by EAuFechaHora "

    
    'Set RsAux = cBase.OpenResultset(Cons) ' , rdOpenForwardOnly, rdConcurReadOnly)
    If ObtenerResultSet(cBase, RsAux, Cons, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
    Do While Not RsAux.EOF
        
        With vsGridArt
            If Not NoRepetido(RsAux("EAuTipo"), RsAux("EAuDocumento")) Then
                
                If RsAux!CliTipo = 1 Then
                    Cons = Trim(RsAux("NomCli"))
                Else
                    Cons = Trim(RsAux("NomEmp"))
                End If
                
                If RsAux("EAuArticulo") = 0 Then
                    If Not IsNull(RsAux("DocNumero")) Then
                        If RsAux("DocNumero") > 0 Then
                            sDoc = "Servicio " & RsAux("SerCodigo") & " fac: " & RsAux("DocSerie") & "-" & RsAux("DocNumero")
                        Else
                            sDoc = "Servicio " & RsAux("SerCodigo")
                        End If
                    Else
                        sDoc = "Servicio " & RsAux("SerCodigo")
                    End If
                Else
                    sDoc = Trim(RsAux("DocSerie")) & "-" & RsAux("DocNumero")
                End If
                
                loc_InsertGridDocumento RsAux("EAuTipo"), RsAux("EAuDocumento"), Cons, sDoc, False, (RsAux("ClienteLlamado") = 1)
                
            End If
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
evExit:
    If bSonido And vsGridArt.Rows > 0 Then loc_SetSonido paSonidoTimbre
    TArr.Enabled = True
    TArr.Interval = 4000

Exit Sub
errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar documentos para entregar.", Err.Description
    TArr.Enabled = True
    TArr.Interval = 4000
End Sub

Private Sub loc_BuscoServicio(ByVal pSerDoc As Long, Optional iTipo As Byte)
On Error GoTo errServicio
Dim sCons As String
Dim RsAux As rdoResultset
Dim bCargue As Boolean

    If Not BuscoSerEnAuxiliar(pSerDoc) Then
        If MsgBox("Este documento no pasó por el Lector, ¿Desea entregar la mercaderia?", vbYesNo + vbDefaultButton1, "Atención") = vbNo Then Exit Sub
    End If
    
    sCons = " Select SerCodigo, SerDocumento, SerCostoFinal, SerEstadoServicio From Servicio "
    
    If iTipo = 0 Then 'es una factura
        sCons = sCons & " Where SerDocumento = " & pSerDoc
    Else
        sCons = sCons & " Where SerCodigo = " & pSerDoc
    End If
    'Set RsAux = cBase.OpenResultset(sCons, rdOpenDynamic, rdConcurValues)
    If ObtenerResultSet(cBase, RsAux, Cons, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
    If Not RsAux.EOF Then
        If IsNull(RsAux("SerCostoFinal")) Then loc_ShowMsg "El costo del servicio no está validado, consulte la ficha del servicio.", 2500, Advierto: RsAux.Close: Exit Sub
        If RsAux("SerCostoFinal") > 0 And IsNull(RsAux("SerDocumento")) Then loc_ShowMsg "No está paga la reparación, pase por caja", 3000, Advierto: RsAux.Close: Exit Sub
        If RsAux("SerEstadoServicio") <> 3 Then loc_ShowMsg "El estado del servicio no es taller, ATENCIÓN!!!", 3000, Advierto: RsAux.Close: Exit Sub
        If Not IsNull(RsAux("SerDocumento")) Then iDocumento = RsAux("SerDocumento") Else iDocumento = RsAux("SerCodigo")
        lCodServicio = RsAux("SerCodigo")
        bCargue = True
    End If
    RsAux.Close
    
    If Not bCargue Then
        loc_ShowMsg "El dato ingresado no es un servicio listo para entregar. ", 3000, Advierto
    Else
        Cons = " Select ArtNombre from Servicio,Producto,Articulo where SerCodigo = " & lCodServicio _
               & " And SerProducto = ProCodigo And ProArticulo = ArtId "
        
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If Not RsAux.EOF Then
            
            vsGridArt.Rows = 0
            With vsGridArt
                .AddItem "Servicio " & lCodServicio
                .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux("ArtNombre"))
                .Cell(flexcpForeColor, .Rows - 1, 0) = vbRed
            End With
        End If
        RsAux.Close
        iPaso = 4
        If iIDUsuario = 0 Then
            lingreso.Caption = "Ingrese su usuario"
        Else
            lingreso.Caption = "Código de servicio"
        End If
        TArr.Enabled = True
        TArr.Interval = 3000
        vsGrid.Rows = 0
    End If
    loc_FocoTBarra
    
Exit Sub
errServicio:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el servicio.", Err.Description
End Sub



'Private Sub BuscoRemito(ByVal Numero As Long)
'On Error GoTo errRemito
'
'    iTipoDoc = 0
'    iDocumento = 0
'
'    Cons = " Select NomCli = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1))+ " _
'           & " RTrim(' ' + CPeNombre2), NomEmp = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')'), CliTipo, " _
'            & " Documento.*, RemDocumento, RemCodigo, RemModificado  " _
'            & " From  Remito, " _
'                    & "Documento Left Outer Join Cliente on Documento.DocCliente = Cliente.CliCodigo " _
'                    & " Left Outer Join Cpersona on Cliente.CliCodigo = CPersona.CpeCliente " _
'                    & " Left Outer Join CEmpresa On CliCodigo = CEmCliente " _
'            & " Where RemDocumento = DocCodigo and RemCodigo = " & Numero
'
'    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
'
'    If Not RsAux.EOF Then
'        If RsAux!DocAnulado Then
'            Screen.MousePointer = 0
'            loc_ShowMsg "El documento ingresado está anulado, consulte en mostrador.", 4000, Advierto
'            RsAux.Close
'            Exit Sub
'        Else
'            If Not IsNull(RsAux!DocPendiente) Then
'                Screen.MousePointer = 0
'                loc_ShowMsg "Documento pendiente de entrega, consulte en mostrador.", 4000, Advierto
'                RsAux.Close
'                Exit Sub
'            End If
'        End If
'    Else
'        Screen.MousePointer = 0
'        iDocumento = 0
'        loc_ShowMsg "No existe un remito para las características ingresadas.", 4000, Advierto
'        RsAux.Close
'        Exit Sub
'    End If
'
'    iDocumento = Numero
'    iTipoDoc = TipoDocumento.Remito
'    gFechaDocumento = RsAux!RemModificado     'Siempre guardo la del Documento
'
'    Dim sNomCli As String, sFactura As String
'
'    If RsAux!CliTipo = 1 Then
'        sNomCli = Trim(RsAux!NomCli)
'    Else
'        sNomCli = Trim(RsAux!NomEmp)
'    End If
'    sFactura = RsAux("DocSerie") & "-" & RsAux("DocNumero")
'    RsAux.Close
'
'    If iDocumento > 0 Then
'
'        Cons = "Select * from entregaAuxiliar where EAuDocumento = " & Numero
'        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
'        If RsAux.EOF Then
'            SinEntregaAuxiliar
'            RsAux.Close
'            Exit Sub
'        End If
'        RsAux.Close
'
'        Cons = " Select ArtId, ArtNombre, ArtTipo, EAuArticulo, EAuCantidad, RReAEntregar ARetirar, ArtNroSerie, Null AEsID, Null AEsNroSerie, Null Desde, Null Hasta " _
'        & " From EntregaAuxiliar Inner join RenglonRemito on  EntregaAuxiliar.EAuDocumento = RenglonRemito.RReRemito " _
'        & " Inner join Articulo on articulo.ArtId = renglonRemito.RReArticulo " _
'        & " Where EAuEstado = 2 And EAuTipo = 6 And RReArticulo = EntregaAuxiliar.EAuArticulo And " _
'        & " RReRemito = " & iDocumento
'
'        CargoArticulos Cons
'
'    End If
'    Screen.MousePointer = 0
'    Exit Sub
'errRemito:
'    clsGeneral.OcurrioError "Error al cargar la información del remito.", Err.Description
'End Sub

Private Sub BuscoRemito(ByVal idRemito As Long)
On Error GoTo errBR
Dim sNomCli As String
    Screen.MousePointer = 11
    iRemito = 0
    Cons = "SELECT NomCli = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1))+ " _
           & " RTrim(' ' + CPeNombre2), NomEmp = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')'), CliTipo, Documento.*, RemCodigo, RemModificado " & _
        " FROM Remito " & _
        "INNER JOIN Documento ON RemDocumento = DocCodigo AND DocAnulado = 0 " & _
        "LEFT OUTER JOIN Cliente ON Documento.DocCliente = Cliente.CliCodigo " & _
        "LEFT OUTER JOIN Cpersona on Cliente.CliCodigo = Cpersona.CpeCliente LEFT OUTER JOIN CEmpresa On CliCodigo = CEmCliente " & _
        "WHERE RemCodigo = " & idRemito
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        
        iTipoDoc = TipoDocumento.Remito
        'gFechaDocumento = RsAux!DocFModificacion
        gFechaDocumento = RsAux!RemModificado
        iDocumento = RsAux!DocCodigo
        iRemito = RsAux("RemCodigo")
        '-------------------------------------------------------------------
        
        If RsAux!CliTipo = 1 Then
            sNomCli = Trim(RsAux!NomCli)
        Else
            sNomCli = Trim(RsAux!NomEmp)
        End If
    
    End If
    RsAux.Close
        
    If iDocumento > 0 And iRemito > 0 Then
        
        Cons = " Select ArtId, ArtNombre ArtNombre, ArtTipo, RReAEntregar EAuCantidad, ArtNroSerie, AEsID, AEsNroSerie, Null Desde, Null Hasta, RReAEntregar ARetirar " _
            & " FROM RenglonRemito INNER JOIN Articulo ON RReArticulo = ArtID " _
            & " LEFT OUTER JOIN ArticuloEspecifico ON AEsArticulo = ArtID And 1 = AEsTipoDocumento And AEsDocumento = " & iDocumento _
            & " WHERE RReRemito = " & iRemito & " AND RReAEntregar > 0"
            
        CargoArticulos Cons
        
    Else
        Screen.MousePointer = 0
        iDocumento = 0
        MsgBox "No existe un remito para las características ingresadas.", vbExclamation, "ATENCIÓN"
    End If
    Screen.MousePointer = 0
    Exit Sub
errBR:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar la información del remito.", Err.Description
End Sub

Private Sub BuscoDocumento(ByVal iTipo As Long, Optional Codigo As Long = 0)
On Error GoTo errBuscoD
Dim gcumplirServicio As Long
Dim sNomCli As String, sFactura As String
Dim idArt As Long, iCliente As Long
    
    iDocumento = 0
           
    Cons = "Select NomCli = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1))+ " _
           & " RTrim(' ' + CPeNombre2), NomEmp = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')'), CliTipo, Documento.* " _
           & " From Documento left outer join Cliente ON Documento.DocCliente = Cliente.CliCodigo " _
           & " Left Outer Join Cpersona on Cliente.CliCodigo = Cpersona.CpeCliente  Left Outer Join CEmpresa " _
           & " On CliCodigo = CEmCliente Where DocCodigo = " & Codigo & " And DocTipo = " & iTipo
           
        
    'Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If ObtenerResultSet(cBase, RsAux, Cons, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
    
    If Not RsAux.EOF Then
        
        If RsAux!DocAnulado Then
            RsAux.Close
            Screen.MousePointer = 0
            loc_ShowMsg "El documento ingresado está anulado, consulte en mostrador.", 4000, Advierto
            Exit Sub
        ElseIf Not IsNull(RsAux!DocPendiente) Then
            RsAux.Close
            Screen.MousePointer = 0
            loc_ShowMsg "Documento pendiente de entrega, consulte en mostrador.", 4000, Advierto
            Exit Sub
        End If
        iDocumento = RsAux!DocCodigo
        sFactura = RsAux("DocSerie") & "-" & RsAux("DocNumero")
    Else
        RsAux.Close
        loc_ShowMsg "No existe un documento para las características ingresadas.", 4000, Advierto
        Exit Sub
    End If
    
    
    iTipoDoc = iTipo
    gFechaDocumento = RsAux!DocFModificacion
    iCliente = RsAux("DocCliente")
    
    If RsAux!CliTipo = 1 Then
        sNomCli = Trim(RsAux!NomCli)
    Else
        sNomCli = Trim(RsAux!NomEmp)
    End If
    
    RsAux.Close
    
    If iDocumento > 0 Then
        Select Case iTipo
            Case TipoDocumento.NotaCredito, TipoDocumento.NotaDevolucion, TipoDocumento.NotaEspecial
                
                Cons = " Select ArtId, ArtNombre ArtNombre, ArtTipo, DevCantidad EAuCantidad, ArtNroSerie, Null AEsID, Null AEsNroSerie, Null Desde, Null Hasta, DevCantidad ARetirar " _
                    & " From Devolucion, Articulo " _
                    & " Where DevCliente = " & iCliente _
                    & " And DevNota = " & iDocumento _
                    & " And DevLocal is Null And DevArticulo = ArtID"
                    
                CargoArticulos Cons
        
        
            Case Else
                Cons = "SELECT * FROM EntregaAuxiliar WHERE EAuDocumento = " & Codigo & " AND EAuEstado IN (1,2) AND EAuLocal = " & paCodigoDeSucursal
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                If Not RsAux.EOF Then
                    idArt = RsAux("EAuArticulo")
                Else
                    SinEntregaAuxiliar
                    RsAux.Close
                    Exit Sub
                End If
                RsAux.Close
                
                If idArt = 0 Then loc_BuscoServicio Codigo, 0: Screen.MousePointer = 0: Exit Sub
                
                Cons = " Select ArtId, IsNull(AEsNombre, ArtNombre) ArtNombre, ArtTipo, EAuCantidad,  RenARetirar ARetirar, ArtNroSerie, AEsID, AEsNroSerie, " _
                    & "(Case IsNull(Convert(int, ArtDisponibleDesde), 0) When 0 Then DocFecha Else (Case DatePart(hh, DocFRetira) When 1 Then DocFRetira - DatePart(n, DocFRetira) Else DocFecha End)" _
                    & " End) Desde, DocFRetira Hasta " _
                    & " From ((((EntregaAuxiliar Inner join Renglon on  EntregaAuxiliar.EAuDocumento = Renglon.RenDocumento) INNER JOIN Documento ON DocCodigo = RenDocumento) " _
                    & " Inner join Articulo on articulo.ArtId = Renglon.RenArticulo) " _
                    & " Left Outer Join ArticuloEspecifico ON AEsArticulo = EAuArticulo And 1 = AEsTipoDocumento And AEsDocumento = EAuDocumento) " _
                    & " Where  Renglon.RenArticulo = EntregaAuxiliar.EAuArticulo and " _
                    & " RenDocumento = " & iDocumento & " and EAuEstado IN(1, 2)"

                CargoArticulos Cons
        End Select
        
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errBuscoD:
    clsGeneral.OcurrioError "Ocurrió un error al buscar un documento.", Err.Description
End Sub

Private Function fnc_RetornoId(ByVal sCodBar As String, ByRef iRetQ As Integer, ByRef iArtEsp As Long, ByRef bEsNSerie As Boolean, ByRef iDoc As Long, ByRef esCombo As Boolean) As Long
On Error GoTo errRetorno
Dim sQy As String
Dim RsAux As rdoResultset
    
    iRetQ = 1
    iArtEsp = 0
    bEsNSerie = False
    
    sCodBar = Replace(sCodBar, "'", "''")
    sQy = "EXEC prg_BuscarArticuloEscaneado '" + sCodBar + "'"
    Set RsAux = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux("ArtID")) Then
            fnc_RetornoId = RsAux("ArtId")
            iArtEsp = RsAux("AEsID")
            iDoc = RsAux("AEsDocumento")
            
            'If Not IsNull(RsAux("ACBCantidad")) Then
            iRetQ = RsAux("ACBCantidad")
            esCombo = RsAux("EsCombo")
            
            If iRetQ > 0 Then
                
                If IsNumeric(sCodBar) Then
                    'Verifico si lo que ingreso es el código del artículo
                    If Trim(sCodBar) = Trim(RsAux("ArtCodigo")) Then
                        
                        Dim iQ As Integer
                        For iQ = 1 To UBound(arrSucesoArt)
                            If arrSucesoArt(iQ) = RsAux("ArtID") Then
                                iQ = 9999
                                Exit For
                            End If
                        Next
                        If iQ <> 9999 Then
                            ReDim Preserve arrSucesoArt(UBound(arrSucesoArt) + 1)
                            arrSucesoArt(UBound(arrSucesoArt)) = RsAux("ArtId")
                        End If
                    End If
                End If
            End If
            If Trim(sCodBar) <> Trim(RsAux("ArtCodigo")) And RsAux("ACBLargo") > 0 Then bEsNSerie = (RsAux("ACBLargo") > 0)
        End If
    End If
    RsAux.Close
    
'    If fnc_RetornoId = 0 Then
'        If bSinEnAuxiliar Then
'            sQy = "Select AEsID, AEsArticulo " & _
'                " From ArticuloEspecifico " & _
'                " Where AEsID = " & sCodBar & " And AEsTipoDocumento = 1 And AEsDocumento = " & iDocumento
'        Else
'            sQy = "Select AEsID, AEsArticulo " & _
'                "From (EntregaAuxiliar Inner Join ArticuloEspecifico ON AEsArticulo = EAuArticulo And 1 = AEsTipoDocumento And AEsDocumento = EAuDocumento) " & _
'                " Where AEsID = " & sCodBar & " And EAuEstado = 2"
'        End If
'        Set RsAux = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
'        If Not RsAux.EOF Then
'            fnc_RetornoId = RsAux(1)
'            iArtEsp = RsAux(0)
'        End If
'        RsAux.Close
'    End If

Exit Function
errRetorno:

End Function

Private Sub CargoArticulos(ByVal sCons As String)
On Error GoTo errCargar
Dim RsAux As rdoResultset
Dim lIdArt As Long
Dim iTipoArt As Integer
Dim iCantidad As Integer

    vsGridArt.Tag = "Art"
    vsGridArt.Rows = 0
    Erase arrArtFecha
    ReDim arrArtFecha(0)
    loc_ShowArts False
    
    'Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If ObtenerResultSet(cBase, RsAux, Cons, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
    If Not RsAux.EOF Then
        vsGridArt.BackColorBkg = &HD0FFFF
        Do While Not RsAux.EOF
            With vsGridArt
            
'                If RsAux("EAuCantidad") > RsAux("ARetirar") Then
'                    RsAux.Close
'                    MsgBox "IMPORTANTE!!!" & vbCrLf & vbCrLf & "Se alteró la cantidad a retirar, debe eliminar el documento de la lista y escanearlo nuevamente.", vbExclamation, "ATENCIÓN"
'                    Exit Sub
'                 End If

                
                If Not IsNull(RsAux("AEsID")) Then
                    .AddItem "E" & RsAux("AEsID") & ":" & RsAux("ArtNombre")
                Else
                    .AddItem RsAux("ArtNombre")
                End If
                 
                 lIdArt = RsAux("ArtId"): .Cell(flexcpData, .Rows - 1, 0) = lIdArt
                 
                 '23/5/2013 Siempre muestro lo pendiente a retirar de la factura.
                 iCantidad = RsAux("ARetirar"): .Cell(flexcpData, .Rows - 1, 1) = iCantidad
                 'iCantidad = RsAux("EAuCantidad"): .Cell(flexcpData, .Rows - 1, 1) = iCantidad
                 .Cell(flexcpText, .Rows - 1, 1) = iCantidad
                 
                iTipoArt = RsAux("ArtTipo")
                
                .Cell(flexcpData, .Rows - 1, 2) = iTipoArt
                
                If RsAux("ArtNroSerie") Then sCons = "s" Else sCons = "n"
                .Cell(flexcpText, .Rows - 1, 2) = sCons
                
                If Not IsNull(RsAux("AEsID")) Then
                    lIdArt = RsAux("AEsID"): .Cell(flexcpData, .Rows - 1, 3) = lIdArt
                    If Not IsNull(RsAux("AEsNroSerie")) Then .Cell(flexcpText, .Rows - 1, 2) = "s" & Trim(RsAux("AEsNroSerie"))
                End If
                
                'Guardo la fecha desde y la fecha hasta.
                If Not IsNull(RsAux("Desde")) And Not IsNull(RsAux("Hasta")) Then
                    ControlFueraDeFecha RsAux("desde"), RsAux("Hasta"), Trim(RsAux("ArtNombre")), RsAux("ArtId")
                    
'                    If Not (CDate(Format(RsAux("Desde"), "dd/mm/yyyy")) <= Date And CDate(Format(RsAux("Hasta"), "dd/mm/yyyy")) >= Date) Then
'
'                        '27/7/2012 como le dan escanear enseguida al usuario se están comiendo este mensaje y dicen que no los cargan en la grilla.
'                        loc_SetSonido "entregamal.wav"
'
'                        If MsgBox("La factura contiene al artículo '" & Trim(RsAux("ArtNombre")) & "' que posee un rango de fechas que no le permite entregarlo." & vbCrLf & vbCrLf & "¿Desea entregarlo de todas formas?", vbQuestion + vbYesNo + vbDefaultButton2, "Artículo con rango de entrega") = vbYes Then
'                            ReDim Preserve arrArtFecha(UBound(arrArtFecha) + 1)
'                            With arrArtFecha(UBound(arrArtFecha))
'                                .Articulo = RsAux("ArtId")
'                                .NroSerie = Trim(RsAux("ArtNombre"))
'                            End With
'                        Else
'                            .RemoveItem .Rows - 1
'                        End If
'                    End If
                    
                End If
                
            End With
        RsAux.MoveNext
        Loop
        RsAux.Close
        
        'Como pude borrar por el rango de fechas pregunto la cantidad de filas que cargue.
        If vsGridArt.Rows = 0 Then
            loc_ShowMsg "No hay artículos para retirar", 4000, Advierto
            CargoFacturas
            Exit Sub
        End If
        
        iPaso = 2
        loc_FocoTBarra
        If iIDUsuario = 0 Then
            lingreso.Caption = "Ingrese su usuario:"
        Else
            lingreso.Caption = "Ingrese un artículo:"
        End If
    Else
        loc_ShowMsg "El documento ingresado no tiene artículos para retirar", 4000, Advierto
        CargoFacturas
    End If
    
Exit Sub
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al Cargar los articulos.", Err.Description
End Sub

Sub ControlFueraDeFecha(ByVal desde As Date, ByVal hasta As Date, ByVal NombreArticulo As String, ByVal IDArticulo As Long)
    
    If Not (CDate(Format(desde, "dd/mm/yyyy")) <= Date And CDate(Format(hasta, "dd/mm/yyyy")) >= Date) Then
        '27/7/2012 como le dan escanear enseguida al usuario se están comiendo este mensaje y dicen que no los cargan en la grilla.
        Dim ret As VbMsgBoxResult
        Do
            loc_SetSonido "entregamal.wav"
            ret = MsgBox("La factura contiene al artículo '" & Trim(NombreArticulo) & "' que posee un rango de fechas que no le permite entregarlo." & vbCrLf & vbCrLf & "¿Desea entregarlo de todas formas?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Artículo con rango de entrega")
        Loop Until ret <> vbCancel
        If ret = vbYes Then
            ReDim Preserve arrArtFecha(UBound(arrArtFecha) + 1)
            ReDim Preserve arrArtFecha(UBound(arrArtFecha) + 1)
            With arrArtFecha(UBound(arrArtFecha))
                .Articulo = IDArticulo
                .NroSerie = Trim(NombreArticulo)
            End With
        Else
            vsGridArt.RemoveItem vsGridArt.Rows - 1
        End If
    End If

End Sub

Private Function ValidoArticuloDeCombo(ByVal oArt As clsArticuloEntrega, ByVal idEsp As Long) As Boolean

    With vsGridArt
        For I = 0 To .Rows - 1
            If oArt.IDArticulo = .Cell(flexcpData, I, 0) And Val(.Cell(flexcpData, I, 3)) = idEsp Then
                If oArt.Cantidad <= .Cell(flexcpText, I, 1) Then
                    ValidoArticuloDeCombo = True
                Else
                    loc_ShowMsg "ATENCIÓN!!! la cantidad de artículos no coincide.", 4000, Advierto
                    tcBarra.SetFocus
                    Exit Function
                End If
            ElseIf oArt.IDArticulo = .Cell(flexcpData, I, 0) Then
                loc_ShowMsg "ATENCIÓN!!! Está entregando un artículo que debería ser específico.", 4000, Advierto
                tcBarra.SetFocus
                Exit Function
            Else
                loc_ShowMsg "ATENCIÓN!!! Está entregando un artículo que no está en la grilla.", 4000, Advierto
                tcBarra.SetFocus
                Exit Function
            End If
        Next
    End With
    
End Function

Private Sub BuscarArt(Texto As String)
On Error GoTo errBuscar
Dim I As Integer
Dim sBarCode As String
Dim iCant As Integer
Dim lIdArt As Long
Dim bEsta As Boolean
    
    If InStr(1, Texto, "*") > 0 Then
        If Not IsNumeric(Mid(Texto, 1, InStr(Texto, "*") - 1)) Then loc_ShowMsg "No existe un artículo para las características ingresadas.", 4000, Advierto: Exit Sub
        sBarCode = (Trim(Mid(Texto, InStr(Texto, "*") + 1, Len(Texto))))
        iCant = CInt(Mid(Texto, 1, InStr(Texto, "*") - 1))
    Else
        iCant = 1
        sBarCode = Texto
    End If
    
    If iCant < 0 Then
        loc_ShowMsg "La cantidad ingresada no es correcta.", 4000, Advierto: Exit Sub
        tcBarra.SetFocus
        Exit Sub
    End If
    
    Dim iArtEsp As Long, bEsNSerie As Boolean, esCombo As Boolean
    Dim iDoc As Long
    lIdArt = fnc_RetornoId(sBarCode, I, iArtEsp, bEsNSerie, iDoc, esCombo)
    If lIdArt = 0 Then
        loc_ShowMsg "No existe un artículo para las características ingresadas.", 4000, Advierto: Exit Sub
    ElseIf iArtEsp > 0 And iDoc <> iDocumento And iDoc > 0 Then
        loc_ShowMsg "El artículo específico no está asociado a este documento. NO ENTREGAR", 4000, Advierto: Exit Sub
    End If
    
    'Si encontré un especifico y no tiene documento entonces no lo considero.
    If iArtEsp > 0 And iDoc = 0 Then iArtEsp = 0
    
    If I = 0 Then I = 1 'x las dudas
    'Le multiplico la cantidad del cód de barras.
    iCant = iCant * I
    
    If esCombo Then
        Dim colCombo As New Collection
        Dim oArtCombo As clsArticuloEntrega
        'Tengo que buscar los ids del combo y si estos requieren Nro. serie pedirlos.
        Cons = "SELECT ParCantidad, PArArticulo, ArtNroSerie, ArtNombre FROM Presupuesto INNER JOIN PresupuestoArticulo ON PreID = PArPresupuesto" & _
            "INNER JOIN Articulo ON ArtId = PArArticulo WHERE PreArtCombo = " & lIdArt
        If ObtenerResultSet(cBase, RsAux, Cons, "comercio") <> RAQ_SinError Then Exit Sub
        Do While Not RsAux.EOF
            'Tengo que recorrer la grilla y buscar el artículo con su cantidad.
            Set oArtCombo = New clsArticuloEntrega
            colCombo.Add oArtCombo
            With oArtCombo
                .Cantidad = RsAux("ParCantidad")
                .IDArticulo = RsAux("PArArticulo")
                .PedirNroSerie = RsAux("ArtNroSerie")
                .NombreArticulo = Trim(RsAux("ArtNombre"))
            End With
            If Not ValidoArticuloDeCombo(oArtCombo, iArtEsp) Then
                RsAux.Close
                Exit Sub
            End If
            RsAux.MoveNext
        Loop
        RsAux.Close
        'Recorro en la grilla a ver si puedo restar los artículos.
        Dim frmPNroSerie As frmNumeroDeSerie
        For Each oArtCombo In colCombo
            'Pido el nro de serie para los que preciso.
            If oArtCombo.PedirNroSerie Then
                Set frmPNroSerie = New frmNumeroDeSerie
                frmPNroSerie.Articulo = oArtCombo
                If frmPNroSerie.Result <> vbOK Then
                    loc_ShowMsg "Combo interrumpido, reinicie escaneo.", 4000, Advierto
                    tcBarra.SetFocus
                    Exit Sub
                End If
            End If
        Next
        For Each oArtCombo In colCombo
        
            With vsGridArt
                For I = 0 To .Rows - 1
                    If oArtCombo.IDArticulo = .Cell(flexcpData, I, 0) Then
                        .Cell(flexcpBackColor, I, 0, , .Cols - 1) = &HC0C0C0
                        .Cell(flexcpText, I, 1) = .Cell(flexcpText, I, 1) - iCant
                        If .Cell(flexcpText, I, 1) = 0 Then .RowHidden(I) = True
                        If oArtCombo.Cantidad = 1 And (Mid(.Cell(flexcpText, I, 2), 1, 1) = "s" Or oArtCombo.PedirNroSerie) Then
                            'Pido nro.
                            If Not arrAgregoElemento(oArtCombo.IDArticulo, sBarCode) Then
                                fnc_PedirSerie oArtCombo.IDArticulo, I
                            End If
                        End If
                        bEsta = True
                        Exit For
                    Else
                        loc_ShowMsg "ATENCIÓN!!! Está entregando de más", 4000, Advierto
                        tcBarra.SetFocus
                        Exit Sub
                    End If
                    Exit For
                 'End If
                Next
                If bEsta Then
                    'Si no pido nro de serie
                    If iCant <> 1 Or .Cell(flexcpText, I, 2) = "n" Or bEsNSerie Then loc_AccionFinalizar
                Else
                    loc_ShowMsg "Ese artículo no está en la grilla, NO ENTREGAR!!!", 4000, Advierto
                End If
            End With

        Next



    End If
    
    bEsta = False
    With vsGridArt
        For I = 0 To .Rows - 1
            If lIdArt = .Cell(flexcpData, I, 0) And Val(.Cell(flexcpData, I, 3)) = iArtEsp Then
                If iCant <= .Cell(flexcpText, I, 1) Then
                    .Cell(flexcpBackColor, I, 0, , .Cols - 1) = &HC0C0C0
                    .Cell(flexcpText, I, 1) = .Cell(flexcpText, I, 1) - iCant
                    If .Cell(flexcpText, I, 1) = 0 Then .RowHidden(I) = True
                    If iCant = 1 And (Mid(.Cell(flexcpText, I, 2), 1, 1) = "s" Or bEsNSerie = True) Then
                        If bEsNSerie Then
                            If Not arrAgregoElemento(lIdArt, sBarCode) Then
                                bEsNSerie = False
                                fnc_PedirSerie lIdArt, I
                            End If
                        Else
                            fnc_PedirSerie lIdArt, I
                        End If
                    End If
                    bEsta = True
                    Exit For
                Else
                    loc_ShowMsg "ATENCIÓN!!! Está entregando de más", 4000, Advierto
                    tcBarra.SetFocus
                    Exit Sub
                End If
                Exit For
             End If
        Next
        If bEsta Then
            'Si no pido nro de serie
            If iCant <> 1 Or .Cell(flexcpText, I, 2) = "n" Or bEsNSerie Then loc_AccionFinalizar
        Else
            loc_ShowMsg "Ese artículo no está en la grilla, NO ENTREGAR!!!", 4000, Advierto
        End If
    End With
    loc_FocoTBarra
    lIdArt = 0
Exit Sub
errBuscar:
    clsGeneral.OcurrioError "Ocurrió un error al buscar los articulos.", Err.Description
End Sub

Private Function loc_BuscarUsuario(Texto As String) As Boolean
On Error GoTo errBuscarCli
Dim Cons As String
Dim RsAux As rdoResultset
    Cons = " Select UsuCodigo, UsuIdentificacion  from Usuario where UsuBarCode = '" & Texto & "'"
    'Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If ObtenerResultSet(cBase, RsAux, Cons, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Function
    If Not RsAux.EOF Then
        iIDUsuario = CLng(RsAux("UsuCodigo"))
        lbUsuario.Caption = Trim(RsAux("UsuIdentificacion"))
        loc_BuscarUsuario = True
    End If
    RsAux.Close
Exit Function
errBuscarCli:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el usuario.", Err.Description
End Function

Private Function fnc_PedirSerie(ByVal idArt As Long, ByVal iRowControl As Integer) As Boolean
On Error Resume Next
    shSerie.Visible = True
    lserie.Visible = True
    vsGridArt.Visible = False
    lserie.Caption = "Ingrese el Nro de serie del Artículo"
    lingreso.Caption = "Ingrese el Nro de serie:"
    lIdArticulo = idArt
    iPaso = 3
    iRowSerie = iRowControl
End Function
Private Function arrAgregoElemento(aIdArticulo As Long, aSerie As String) As Boolean
    
    On Error GoTo errAgregar
    arrAgregoElemento = False
    If arrBuscoElemento(aIdArticulo, aSerie) <> 0 Then
        loc_ShowMsg "El nro. de serie ingresado ya fue entregado !!!!." & vbCrLf & "Nº Serie: " & Trim(aSerie), 5000, Advierto
        Exit Function
    End If


    'Antes de agregarlo verifico si es especifico y tiene el mismo nro. de serie.
    If Val(vsGridArt.Cell(flexcpData, iRowSerie, 3)) > 0 And Len(vsGridArt.Cell(flexcpText, iRowSerie, 2)) > 1 Then
        If LCase(Trim(Mid(vsGridArt.Cell(flexcpText, iRowSerie, 2), 2))) <> LCase(Trim(aSerie)) Then
            loc_ShowMsg "Nro. de Serie INCONRRECTO, está entregando un ARTÍCULO ESPECÍFICO", 5000, Error
            Exit Function
        End If
    End If

    Dim aIdxC As Integer
    aIdxC = UBound(arrNroSerie) + 1
    ReDim Preserve arrNroSerie(aIdxC)
        
    arrNroSerie(aIdxC).Articulo = aIdArticulo
    arrNroSerie(aIdxC).NroSerie = Trim(aSerie)
    
    arrAgregoElemento = True
    
Exit Function
errAgregar:
    clsGeneral.OcurrioError "Ocurrió un error al agregar el elemento.", Err.Description
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
Exit Function
errB:
End Function

Private Function AccionGrabar(ByVal iIDUserSuceso As Long, ByVal sDefensa As String, ByVal iUserSucesoRango As Long, ByVal sdefsucrango As String, ByVal sDetSuceso As String) As Boolean
    
    FechaDelServidor
    On Error GoTo errorBT
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
    On Error GoTo errorET
    
    If iTipoDoc <> 6 Then
        'Antes veo si la fecha es la misma-------------------------------------------------------------------------------------
        Cons = "Select * from Documento Where DocCodigo = " & iDocumento
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux!DocFModificacion <> gFechaDocumento Then
            RsAux.Close
            cBase.RollbackTrans
            loc_ShowMsg "El documento seleccionado ha sido modificado por otra terminal. Vuelva a consultar.", 3000, Error
            GoTo errorET
            Exit Function
        Else
            RsAux.Edit
            RsAux!DocFModificacion = Format(gFechaServidor, "yyyy/mm/dd hh:mm:ss")
            RsAux.Update
        End If
        RsAux.Close '-----------------------------------------------------------------------------------------------------------
    Else
        Cons = "SELECT * FROM Remito WHERE RemCodigo = " & iRemito
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux!RemModificado <> gFechaDocumento Then
            RsAux.Close
            cBase.RollbackTrans
            loc_ShowMsg "El documento seleccionado ha sido modificado por otra terminal. Vuelva a consultar.", 3000, Error
            GoTo errorET
            Exit Function
        Else
            RsAux.Edit
            RsAux!RemModificado = Format(gFechaServidor, "yyyy/mm/dd hh:mm:ss")
            RsAux.Update
        End If
        RsAux.Close '-----------------------------------------------------------------------------------------------------------
    End If
'
    If fnc_EsDevolucion Then
        GraboDatosTablasDevolucion
        
        Dim idFactura As Long
        Cons = "Select * from Nota Where NotNota = " & iDocumento
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then idFactura = RsAux!NotFactura
        'Si es cero --> nota especial
        If idFactura > 0 Then BorroProductosVendidos idFactura
        
    Else
        'BorroProductosVendidos 'Por duplicaicones de códigos
        GraboDatosTablas
        GraboProductosVendidos iDocumento
    End If
    
    If iIDUserSuceso > 0 Then
        Dim iQ As Integer
        For iQ = 1 To UBound(arrSucesoArt)
            clsGeneral.RegistroSuceso cBase, Now, 24, paCodigoDeTerminal, iIDUserSuceso, iDocumento, arrSucesoArt(iQ), "Ingresó código del artículo y existe código de barras.", sDefensa
        Next
    End If
    If iUserSucesoRango > 0 Then
        clsGeneral.RegistroSuceso cBase, Now, 25, paCodigoDeTerminal, iUserSucesoRango, iDocumento, 0, "Arts fuera fecha: " + sDetSuceso, sdefsucrango
    End If

    cBase.CommitTrans    'Fin de la TRANSACCION------------------------------------------
    AccionGrabar = True
    
    On Error Resume Next
    oHub.InvokeMethod prmHUBMetod
    
    Screen.MousePointer = 0
    Exit Function

errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación."
    Exit Function
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación.", 4000, Error
    Exit Function

End Function

Private Sub GraboDatosTablasDevolucion()
Dim aDocumento As Long
Dim rsDev As rdoResultset
Dim iQ As Currency

    With vsGridArt
        For I = 0 To .Rows - 1
            If .RowHidden(I) Or (Val(.Cell(flexcpData, I, 1)) <> Val(.Cell(flexcpText, I, 1))) Then

                iQ = (Val(.Cell(flexcpData, I, 1)) - Val(.Cell(flexcpText, I, 1)))
            
                'Actualizo los datos en tabla Devoluciones------------------------------------------------------------------------------------
                Cons = "Select * From Devolucion" & _
                        " Where DevNota = " & iDocumento & _
                        " And DevArticulo = " & .Cell(flexcpData, I, 0) & _
                        " And DevLocal is Null"
                Set rsDev = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                rsDev.Edit
                rsDev!DevLocal = paCodigoDeSucursal
                rsDev!DevFAltaLocal = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
                rsDev.Update
                rsDev.Close
                '-------------------------------------------------------------------------------------------------------------------------------
                
                'Marco el ALTA del STOCK AL LOCAL
                'Genero Movimiento
                MarcoMovimientoStockFisico iIDUsuario, TipoLocal.Deposito, paCodigoDeSucursal, CLng(.Cell(flexcpData, I, 0)), iQ, paEstadoArticuloEntrega, 1, iTipoDoc, iDocumento
                'Alta del Stock en Local
                MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, CLng(.Cell(flexcpData, I, 0)), iQ, paEstadoArticuloEntrega, 1
                
                'Sumo al Stock Total
                MarcoMovimientoStockTotal CLng(.Cell(flexcpData, I, 0)), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, iQ, 1
            End If
        Next
    End With

End Sub

Private Sub GraboProductosVendidos(ByVal iDocumento As Long)
Dim Cons As String
Dim rsPV As rdoResultset
Dim I As Integer

    Cons = "Select * from ProductosVendidos Where PVeDocumento = " & iDocumento
    Set rsPV = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    For I = LBound(arrNroSerie) To UBound(arrNroSerie)
        If arrNroSerie(I).Articulo <> -1 And Trim(arrNroSerie(I).NroSerie) <> "" Then
            rsPV.AddNew
            rsPV!PVeDocumento = iDocumento
            rsPV!PVeArticulo = arrNroSerie(I).Articulo
            rsPV!PVeNSerie = Trim(arrNroSerie(I).NroSerie)
            rsPV.Update
        End If
    Next
    rsPV.Close

End Sub

Private Sub GraboDatosTablas()
Dim iQ As Currency
Dim I As Integer
Dim Cons As String
Dim RsAux As rdoResultset
Dim iEstado As Byte

    'Cambio estado a entregado
    If Not bSinEnAuxiliar Then iEstado = 3 Else iEstado = 0

    With vsGridArt
        For I = 0 To .Rows - 1
            If .RowHidden(I) Or (Val(.Cell(flexcpData, I, 1)) <> Val(.Cell(flexcpText, I, 1))) Then
                iQ = (Val(.Cell(flexcpData, I, 1)) - Val(.Cell(flexcpText, I, 1)))
                If iTipoDoc = TipoDocumento.Remito Then
                
                    Cons = "SELECT * FROM RenglonRemito " _
                        & " WHERE RReRemito = " & iRemito _
                        & " AND RReArticulo = " & .Cell(flexcpData, I, 0)
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If RsAux("RReCantidad") < RsAux("RReAEntregar") Then
                        RsAux.Close
                        RsAux.Edit
                    End If
                    If RsAux("RReAEntregar") < iQ Then
                        'loc_ShowMsg "La cantidad a retirar es mayor a la que puede retirar ", 5000, Advierto: loc_FocoTBarra: CargoFacturas: Exit Sub
                        RsAux.Close
                        RsAux.Edit
                    Else
                        RsAux.Edit
                        RsAux("RReAEntregar") = RsAux("RReAEntregar") - iQ
                        RsAux.Update
                    End If
                    RsAux.Close
                
                Else
                    
                    Cons = "SELECT RenARetirar, RenCantidad FROM Renglon where RenArticulo = " & .Cell(flexcpData, I, 0) _
                         & " AND RenDocumento = " & iDocumento
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If RsAux("RenCantidad") < RsAux("RenAretirar") Then
                        RsAux.Close
                        RsAux.Edit
                    End If
                    
                    If RsAux("RenAretirar") < iQ Then
                        'loc_ShowMsg "La cantidad a retirar es mayor a la que puede retirar ", 5000, Advierto: loc_FocoTBarra: CargoFacturas: Exit Sub
                        RsAux.Close
                        RsAux.Edit
                    Else
                        RsAux.Edit
                        RsAux("RenAretirar") = RsAux("RenAretirar") - iQ
                        RsAux.Update
                    End If
                    RsAux.Close
                    '-------------------------------------------------------------------------------------------------------------------------------
                End If
                If Val(.Cell(flexcpData, I, 2)) <> paTipoArticuloServicio Then
                        cBase.Execute "EXEC prg_EntregaMercaderia_EntregoArticulo " & iIDUsuario & ", " & .Cell(flexcpData, I, 0) & ", " & _
                            iQ * -1 & ", " & iTipoDoc & ", " & IIf(iTipoDoc = TipoDocumento.Remito, iRemito, iDocumento) & ", " & paCodigoDeSucursal & ", " & _
                            paCodigoDeTerminal & ", " & iEstado
                    End If
            Else
                'Lo anulo
                If Not bSinEnAuxiliar Then CambioAEntregado .Cell(flexcpData, I, 0), 4
            End If
        Next
    End With
End Sub

Public Sub CambioAEntregado(ByVal idArt As Long, Optional iEstado As Byte = 3)
Dim Cons As String
Dim RsAux As rdoResultset
    Cons = "Select * from EntregaAuxiliar where EauDocumento = " & iDocumento _
           & " and EAuArticulo = " & idArt & " and EAuestado = 2"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        RsAux.Edit
        RsAux("EAuEstado") = iEstado
        If Abs(DateDiff("s", RsAux("EAuFechaHora"), Now)) > 32000 Then
            RsAux("EAuTiempoTotal") = 380
        Else
            RsAux("EAuTiempoTotal") = Abs(DateDiff("s", RsAux("EAuFechaHora"), Now))
        End If
        RsAux.Update
    End If
    RsAux.Close
End Sub

Private Sub CambiarEstado(ByVal CodSer As Long)
On Error GoTo errCambiarE
Dim Cons As String
Dim RsAux As rdoResultset
    
    FechaDelServidor
    
    On Error GoTo errBT
    cBase.BeginTrans
    On Error GoTo errRB
    
    Cons = "Select SerEstadoServicio, SerFCumplido from servicio where SerCodigo = " & CodSer
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        RsAux.Edit
        RsAux("SerEstadoServicio") = 5
        RsAux("SerFCumplido") = Format(gFechaServidor, "yyyy/mm/dd hh:mm:ss")
        RsAux.Update
    End If
    RsAux.Close
    
    Cons = "Select EauEstado from EntregaAuxiliar where EauDocumento =" & CodSer & " and EAUTipo = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        RsAux.Edit
        RsAux("EAuEstado") = 3
        RsAux.Update
    Else
        RsAux.Close
        Cons = "Select EauEstado from EntregaAuxiliar where EauDocumento =" & iDocumento & " and EAUTipo In (1,2) "
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            RsAux.Edit
            RsAux("EAuEstado") = 3
            RsAux.Update
        End If
        
        
    End If
    RsAux.Close
    
    cBase.CommitTrans
    On Error Resume Next
    
    iPaso = 1
    iIDUsuario = 0
    lingreso.Caption = "Ingrese su usuario"
    loc_ShowMsg "Servicio entregado", 2000, informo
    
Exit Sub
errCambiarE:
    clsGeneral.OcurrioError "Ocurrió un error al cambiar el estado.", Err.Description
    
    
errBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación."
    Exit Sub
errRB:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación.", 4000, Error
    Exit Sub
End Sub

Private Function SinEntregaAuxiliar() As Boolean
On Error GoTo errSEA
Dim sCons As String
Dim RsAux As rdoResultset
Dim lIdArt As Long
Dim iTipoArt As Integer
Dim iCantidad As Integer

    ReDim arrArtFecha(0)
    
    If MsgBox("Este documento no pasó por el Lector, ¿Desea entregar la mercaderia?", vbYesNo + vbQuestion, "Atención") = vbNo Then Exit Function
    'controlar el cancelar
    
    Select Case iTipoDoc
            
        'contado
        Case 1, 2, 6
            sCons = " Select ArtId, IsNull(AEsNombre, ArtNombre) ArtNombre,  RenARetirar ARetirar, ArtNroSerie, AEsID, AEsNroSerie, " _
                & "(Case IsNull(Convert(int, ArtDisponibleDesde), 0) When 0 Then DocFecha Else (Case DatePart(hh, DocFRetira) When 1 Then DocFRetira - DatePart(n, DocFRetira) Else DocFecha End)" _
                    & " End) Desde, DocFRetira + " & lToleranciaEntrega & " Hasta " _
                & " From ((Renglon Inner join documento on DocCodigo = RenDocumento) Inner join Articulo on renglon.RenArticulo = articulo.ArtId) " _
                & " Left Outer Join ArticuloEspecifico ON AEsArticulo = RenArticulo And AEsDocumento = RenDocumento And AEsTipoDocumento = 1 " _
                & " WHERE RenDocumento = " & iDocumento & " and RenARetirar > 0 and artTipo <> " & paTipoArticuloServicio
            
        'remito
'        Case 6
'            sCons = " Select ArtId, ArtNombre, RReAEntregar ARetirar, ArtNroSerie, Null as AEsID, Null as desde, Null as Hasta " _
'                & " From RenglonRemito Inner join Articulo On renglonRemito.RReArticulo = articulo.ArtId " _
'                & " Where RReRemito = " & iDocumento & " and RReAEntregar > 0 and artTipo <> " & paTipoArticuloServicio
            
     End Select
     
     Set RsAux = cBase.OpenResultset(sCons, rdOpenDynamic, rdConcurValues)
     
     If Not RsAux.EOF Then
        vsGridArt.BackColorBkg = &HF0F0F0      'cambiar color de fondo
        vsGridArt.Rows = 0
        loc_ShowArts False
        Do While Not RsAux.EOF
            
            With vsGridArt
                If Not IsNull(RsAux("AEsID")) Then
                    .AddItem "E" & RsAux("AEsID") & ":" & RsAux("ArtNombre")
                Else
                    .AddItem RsAux("ArtNombre")
                End If
                lIdArt = RsAux("ArtId"): .Cell(flexcpData, .Rows - 1, 0) = lIdArt
                iCantidad = RsAux("ARetirar"): .Cell(flexcpData, .Rows - 1, 1) = iCantidad
                .Cell(flexcpText, .Rows - 1, 1) = RsAux("ARetirar")
                If RsAux("ArtNroSerie") Then sCons = "s" Else sCons = "n"
                .Cell(flexcpText, .Rows - 1, 2) = sCons
                
                If Not IsNull(RsAux("AEsID")) Then
                    lIdArt = RsAux("AEsID"): .Cell(flexcpData, .Rows - 1, 3) = lIdArt
                    If Not IsNull(RsAux("AEsNroSerie")) Then .Cell(flexcpText, .Rows - 1, 2) = "s" & Trim(RsAux("AEsNroSerie"))
                End If
                
                'Guardo la fecha desde y la fecha hasta.
                If Not IsNull(RsAux("Desde")) And Not IsNull(RsAux("Hasta")) Then
                    ControlFueraDeFecha RsAux("desde"), RsAux("Hasta"), Trim(RsAux("ArtNombre")), RsAux("ArtId")
                End If
                
            End With
        RsAux.MoveNext
        Loop
        RsAux.Close
        iPaso = 2
        'pongo la variable en true para marcar que pase por aca
        bSinEnAuxiliar = True
        loc_FocoTBarra
        If iIDUsuario = 0 Then
            lingreso.Caption = "Ingrese su Usuario:"
        Else
            lingreso.Caption = "Ingrese un artículo:"
        End If
    Else
        RsAux.Close
        MsgBox "No hay artículos pendientes para el documento, verifique.", vbInformation, "Atención"
    End If
    Exit Function
errSEA:
    clsGeneral.OcurrioError "Error en la busqueda sin entregaAuxiliar.", 4000, Error
End Function

Private Function BuscoSerEnAuxiliar(ByVal lCodSer As Long) As Boolean
On Error GoTo errBSA
Dim sCons As String
Dim RsAux As rdoResultset
    
    BuscoSerEnAuxiliar = False
    
    sCons = "Select * from EntregaAuxiliar where EAuDocumento = " & lCodSer
    Set RsAux = cBase.OpenResultset(sCons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        BuscoSerEnAuxiliar = True
    End If
    RsAux.Close

Exit Function
errBSA:
    clsGeneral.OcurrioError "Error en la busqueda del servicio sin entregaAuxiliar", 4000, Error
End Function

Private Function ControloUsuario(Texto As String) As Boolean
On Error Resume Next
Dim iDBarCode As String
    
    iDBarCode = CStr(Trim(Mid(tcBarra.Text, 2)))
    If Not loc_BuscarUsuario(iDBarCode) Then loc_ShowMsg "El usuario ingresado no existe.", 3000, Advierto: loc_FocoTBarra: Exit Function
    ControloUsuario = True
    
End Function

Private Sub CambiarEstadoArt(ByVal iEstado As Byte, ByVal iDoc As Long, ByVal iTipo As Byte)
On Error GoTo errCEArt
Dim Cons As String
Dim RsAux As rdoResultset

    Cons = "Select EAuEstado from EntregaAuxiliar where EauDocumento = " & iDoc & " And EAuTipo = " & iTipo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        Do While Not RsAux.EOF
            RsAux.Edit
            RsAux("EAuEstado") = iEstado
            RsAux.Update
            RsAux.MoveNext
        Loop
    End If
    RsAux.Close
    loc_ShowMsg "Cambió el estado", 3000, informo
    
    On Error Resume Next
    oHub.InvokeMethod prmHUBMetod
    
    CargoFacturas
    
    
    Exit Sub
errCEArt:
End Sub

Private Sub vsGrid_DblClick()
On Error Resume Next
    If Not tcBarra.Visible Then ps_OcultoDocs Val(vsGrid.Cell(flexcpData, vsGrid.Row, 0)), Val(vsGrid.Cell(flexcpData, vsGrid.Row, 1))
End Sub

Private Sub vsGrid_GotFocus()
On Error Resume Next
    vsGrid_SelChange
End Sub

Private Sub vsGrid_SelChange()
On Error Resume Next
    If vsGrid.Rows > 0 And paArrimar = 0 Then vsGrid.BackColorSel = vsGrid.Cell(flexcpBackColor, vsGrid.Row, vsGrid.Col)
End Sub

Private Sub vsGridArt_DblClick()
On Error Resume Next
    If Not tcBarra.Visible Then ps_OcultoDocs Val(vsGridArt.Cell(flexcpData, vsGridArt.Row, 0)), Val(vsGridArt.Cell(flexcpData, vsGridArt.Row, 1))
End Sub

Private Sub vsGridArt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vsGridArt.Rows > 0 And Button = 2 And iPaso = 1 Then PopupMenu MnuGrilla
End Sub

Private Function BorroProductosVendidos(Optional ByVal iCodDoc As Long = 0)
Dim I As Byte
    
    Dim rsPV As rdoResultset
    For I = LBound(arrNroSerie) To UBound(arrNroSerie)
        If arrNroSerie(I).Articulo <> -1 And Trim(arrNroSerie(I).NroSerie) <> "" Then
            
            Cons = "Select * from ProductosVendidos " & _
                    " Where PVeArticulo = " & arrNroSerie(I).Articulo & _
                    " And PVeNSerie = '" & Replace(Trim(arrNroSerie(I).NroSerie), "'", "") & "'"
            
            If iCodDoc > 0 Then Cons = Cons & " AND PVeDocumento  = " & iCodDoc
            
            Set rsPV = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rsPV.EOF Then rsPV.Delete
            rsPV.Close
        End If
    Next
Exit Function
End Function


Private Sub vsGridArt_SelChange()
    If vsGridArt.Rows > 0 And paArrimar = 0 Then vsGridArt.BackColorSel = vsGridArt.Cell(flexcpBackColor, vsGridArt.Row, vsGridArt.Col)
End Sub
