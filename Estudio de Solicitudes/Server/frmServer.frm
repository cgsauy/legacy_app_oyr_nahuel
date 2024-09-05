VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmServer 
   Caption         =   "Servidor de Asuntos"
   ClientHeight    =   5625
   ClientLeft      =   2805
   ClientTop       =   3450
   ClientWidth     =   6690
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   6690
   Begin VB.Timer tmT 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   540
      Top             =   4800
   End
   Begin VB.PictureBox picImagen 
      BorderStyle     =   0  'None
      Height          =   2355
      Index           =   1
      Left            =   2700
      ScaleHeight     =   2355
      ScaleWidth      =   2775
      TabIndex        =   3
      Top             =   2580
      Width           =   2775
      Begin VB.PictureBox picIcono 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   120
         Picture         =   "frmServer.frx":030A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   8
         Top             =   60
         Width           =   480
      End
      Begin VSFlex6DAOCtl.vsFlexGrid vsError 
         Height          =   1575
         Left            =   60
         TabIndex        =   4
         Top             =   720
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   2778
         _ConvInfo       =   1
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
      Begin VB.Label labSeparador1 
         BorderStyle     =   1  'Fixed Single
         Height          =   45
         Left            =   0
         TabIndex        =   10
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label Label3 
         Caption         =   "Errores/Info"
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
         Left            =   840
         TabIndex        =   9
         Top             =   180
         Width           =   1935
      End
   End
   Begin VB.PictureBox picImagen 
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   0
      Left            =   660
      ScaleHeight     =   3015
      ScaleWidth      =   4635
      TabIndex        =   1
      Top             =   1020
      Width           =   4635
      Begin VB.PictureBox picIcono 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   0
         Left            =   120
         Picture         =   "frmServer.frx":0BD4
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   5
         Top             =   60
         Width           =   480
      End
      Begin VSFlex6DAOCtl.vsFlexGrid vsConexion 
         Height          =   2235
         Left            =   60
         TabIndex        =   2
         Top             =   720
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   3942
         _ConvInfo       =   1
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
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
      Begin VB.Label labSeparador 
         BorderStyle     =   1  'Fixed Single
         Height          =   45
         Left            =   60
         TabIndex        =   7
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Conexiones Activas"
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
         Left            =   840
         TabIndex        =   6
         Top             =   180
         Width           =   3255
      End
   End
   Begin ComctlLib.TabStrip tabServer 
      Height          =   1635
      Left            =   720
      TabIndex        =   0
      Top             =   180
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2884
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Conexiones"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Errores/Info"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock wsSocket 
      Index           =   0
      Left            =   1680
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type typConect
    IdSocket As Long
    IdUsuario As Long
    IdsBooleans As String
End Type

Dim arrConect() As typConect

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    
    InicializoGrillas
        
    wsSocket(0).LocalPort = 0 'prmLocalPort  '63259
    wsSocket(0).Listen
    
    '20070906 Por error de que el puerto se estaba usando, acutalizo prm automáticamente ------------------
    prmLocalPort = wsSocket(0).LocalPort
'    Dim mSQL As String
'    mSQL = "Update Parametro set ParValor = " & prmLocalPort & " Where ParNombre = 'serverasuntos_port_ip'"
'    cBase.Execute mSQL
'    '------------------------------------------------------------------------------------------------------
    
    tmT.Interval = prmQueryInterval
    tmT.Enabled = True
    Exit Sub
    
errLoad:
    MsgBox "Error al iniciar el formulario." & vbCrLf & _
                "Error: " & Trim(Err.Description), vbCritical, "Error"
End Sub

Private Sub InicializoGrillas()

    With vsConexion
        .Rows = 1
        .Cols = 4
        .FormatString = ">ID|<Dirección IP|<Puerto (R)|<S.State|<Usuario|<Sol Ser Gas Suc SRe"
        .ColWidth(1) = 1500
        .ColWidth(4) = 1200
    End With
    
    With vsError
        .Rows = 1: .Cols = 2
        .FormatString = "<Hora|<Error"
        .ColWidth(0) = 780
    End With
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With tabServer
        .Left = 60
        .Top = 60
        .Width = Me.ScaleWidth - (.Left * 2)
        .Height = Me.ScaleHeight - (.Top * 2)
    End With
    
    picImagen(0).ZOrder 0
    
    Dim I As Integer
    For I = 0 To picImagen.Count - 1
        With picImagen(I)
            .Left = tabServer.ClientLeft
            .Width = tabServer.ClientWidth
            .Height = tabServer.ClientHeight
            .Top = tabServer.ClientTop
        End With
    Next
    
    With vsConexion
        .Left = 60
        .Width = picImagen(0).Width - (.Left * 2)
        .Height = picImagen(0).Height - .Top - 60
        .Appearance = flexFlat
    End With
    labSeparador.Width = vsConexion.Width
    With vsError
        .Left = 60
        .Width = picImagen(0).Width - (.Left * 2)
        .Height = picImagen(0).Height - .Top - 60
        .Appearance = flexFlat
    End With
    
    labSeparador.Left = vsError.Left
    labSeparador1.Left = vsError.Left
    labSeparador1.Width = vsError.Width
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Dim Socket As Winsock
    For Each Socket In wsSocket
        ws_Desconecto Socket
    Next
    
    EndMain
    
End Sub

Private Sub Label1_DblClick()
    ReInicializar
End Sub

Private Sub tabServer_Click()
    picImagen(tabServer.SelectedItem.Index - 1).ZOrder 0
End Sub

Private Sub tmT_Timer()
    
On Error GoTo errTimer
Dim mEval As String

    'Debug.Print Time
    If vsConexion.Rows = 1 Then Exit Sub
    tmT.Enabled = False
    
    mEval = "1- Chek Usrs"
    loc_CargoUsuariosLogs
    
    mEval = "2- Ini Tramas"
    arrInicializoTramas
    If idxTrama = -1 Then GoTo etqSalir
    
    Dim I As Integer
    For I = LBound(arrConect) To UBound(arrConect)
        ws_EnivarTramasAlCliente arrConect(I).IdSocket, arrConect(I).IdUsuario, arrConect(I).IdsBooleans
    Next

etqSalir:
    tmT.Enabled = True
    Exit Sub
    
errTimer:
    loc_InsertoError "Evento Timer " & mEval
    ReInicializar
    tmT.Enabled = True
End Sub

Private Function ws_EnivarTramasAlCliente(wsIDSocket As Long, wsIDUsuario As Long, wsBooleans As String) As Boolean
On Error GoTo errSendTrama
Dim bSend As Boolean
Dim idxT As Integer

    If wsSocket(wsIDSocket).State = 7 Then
    
        wsSocket(wsIDSocket).SendData "START" & vbCrLf
        DoEvents
        
        For idxT = LBound(arrTramas) To UBound(arrTramas)
            
            If arrTramas(idxT).IDUserPara = 0 Or arrTramas(idxT).IDUserPara = wsIDUsuario Then
            
                bSend = False
            
                Select Case arrTramas(idxT).IDTipo
                    Case Asuntos.Solicitudes
                            bSend = (wsBooleans Like "1*")
                
                    Case Asuntos.Servicios
                            bSend = (wsBooleans Like "?1*")
                
                    Case Asuntos.GastosAAutorizar
                            bSend = (wsBooleans Like "??1*")
                            
                    Case Asuntos.SucesosAAutorizar
                            bSend = (wsBooleans Like "???1*")
                    
                    Case Asuntos.SolicitudesResueltas
                            bSend = (wsBooleans Like "????1*")
                End Select
                
                If bSend Then
            
                    wsSocket(wsIDSocket).SendData arrTramas(idxT).IDUserPara & "|" & arrTramas(idxT).IDTipo & "|" & _
                                                          arrTramas(idxT).IDEstadoTrama & "|" & arrTramas(idxT).DatosTrama & vbCrLf
                    DoEvents
                
                End If
            End If
            
        Next
        
        wsSocket(wsIDSocket).SendData "END" & vbCrLf
        DoEvents
    
    End If
    ws_EnivarTramasAlCliente = True
    
    Exit Function
    
errSendTrama:
    loc_InsertoError "ws_EnivarTramasAlCliente IDS:" & wsIDSocket
    ws_EnivarTramasAlCliente = False
End Function

Private Sub vsConexion_DblClick()
On Error Resume Next
    If vsConexion.Row >= vsConexion.FixedRows Then
        vsConexion.Cell(flexcpText, vsConexion.Row, 3) = wsSocket(vsConexion.Cell(flexcpValue, vsConexion.Row, 0)).State
    End If
    
End Sub

Private Sub wsSocket_Close(Index As Integer)
    ws_CierroConexion Index
End Sub

Private Sub wsSocket_ConnectionRequest(Index As Integer, ByVal requestID As Long)

Dim mNewIndice As Long
    
    'Siempre escucho con el control cero y le asigno la conexión a un nuevo control.
    If Index = 0 Then
    
        'Retorno el primer socket libre o asigno uno nuevo al array.
        mNewIndice = ws_InstanciaSocket(wsSocket)
        
        If mNewIndice > 0 Then
            wsSocket(mNewIndice).LocalPort = 0      'Asigna el primer puerto disponible.
            wsSocket(mNewIndice).Accept requestID
            
            loc_NewConexion mNewIndice
        Else
        
            loc_InsertoError "ConnectionRequest "
        End If
    End If
    
End Sub

Private Sub wsSocket_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    Dim strDato As String

    If wsSocket(Index).BytesReceived > 0 Then
        
        wsSocket(Index).GetData strDato
        
        loc_ProcesoTrama Index, strDato
                
    End If
    
End Sub

Private Function loc_ProcesoTrama(Index As Integer, mData As String)

Dim Idx As Integer, Idx2 As Integer

    With vsConexion
        For Idx = .FixedRows To .Rows - 1
            If .Cell(flexcpValue, Idx, 0) = Index Then
                'En esta trama Viene el IDUsr|NombreUsr|000 (Info SGS)
                
                '>ID|<Dirección IP|<Puerto (R)|<Puerto (L)|<Usuario|<Info SGS"
                Dim mValue As Long
                Dim arrData() As String
                arrData = Split(mData, "|")
                
                Debug.Print mData
                For Idx2 = LBound(arrData) To UBound(arrData)
                    Select Case Idx2
                        Case 0: mValue = arrData(Idx2): .Cell(flexcpData, Idx, 4) = mValue
                        
                        Case 1: .Cell(flexcpText, Idx, 4) = Trim(arrData(Idx2))
                        Case 2: .Cell(flexcpText, Idx, 5) = Trim(arrData(Idx2))
                    End Select
                Next
                
                Exit For
            End If
        Next
    End With
    
End Function

Private Sub loc_NewConexion(ByVal Index As Long)

    With vsConexion
        .AddItem Index
        .Cell(flexcpText, .Rows - 1, 1) = wsSocket(Index).RemoteHostIP
        .Cell(flexcpText, .Rows - 1, 2) = wsSocket(Index).RemotePort
        .Cell(flexcpText, .Rows - 1, 3) = wsSocket(Index).State   '.LocalPort
    End With
    
End Sub

Private Sub loc_DeleteConexion(ByVal Index As Long)
Dim Idx As Integer
    On Error Resume Next
    With vsConexion
        For Idx = .FixedRows To .Rows - 1
            If .Cell(flexcpValue, Idx, 0) = Index Then
                .RemoveItem Idx
                Exit For
            End If
        Next
    End With
End Sub

Private Sub ws_CierroConexion(ByVal Index As Long)
On Error Resume Next
    'Cierro conexión y elimino de la lista.
    loc_DeleteConexion Index
    ws_Desconecto wsSocket(Index)
    Unload wsSocket(Index)
End Sub

Public Sub loc_InsertoError(ByVal strMsg As String, Optional bInfo As Boolean = False)
    With vsError
        If .Rows > 30 Then .Rows = .FixedRows
        .AddItem Time
        
        If Not bInfo Then
            .Cell(flexcpText, .Rows - 1, 1) = strMsg & " " & Err.Number & " - " & Err.Description
        Else
            .Cell(flexcpText, .Rows - 1, 1) = strMsg
        End If
        
    End With
End Sub

Private Function loc_CargoUsuariosLogs()
Dim I As Integer, mQRows As Integer
Dim iArr As Integer
    iArr = -1
    ReDim arrConect(0)
    prmUserLogs = ""
    mQRows = vsConexion.Rows - 1
    With vsConexion
        For I = .FixedRows To mQRows
            iArr = iArr + 1
            ReDim Preserve arrConect(iArr)
            arrConect(iArr).IdUsuario = .Cell(flexcpData, I, 4)
            arrConect(iArr).IdSocket = .Cell(flexcpValue, I, 0)
            arrConect(iArr).IdsBooleans = Trim(.Cell(flexcpText, I, 5))     'Sol Ser Gas Suc SRe
            If arrConect(iArr).IdUsuario <> 0 Then
                prmUserLogs = prmUserLogs & arrConect(iArr).IdUsuario & ","
            End If
        Next
    End With
    If Trim(prmUserLogs) <> "" Then prmUserLogs = Mid(prmUserLogs, 1, Len(prmUserLogs) - 1)
End Function

Private Function ws_InstanciaSocket(Sockets As Variant) As Long
Dim Indice As Long
    
    On Error GoTo InicioControl
    ws_CheckSocketsStates
    
    For Indice = 1 To 10000
        If Sockets(Indice).Name = "" Then       'Si tengo uno vacío da error
        End If
    Next Indice
    
    'Si no dio error tengo todos asignados entonces le sumo uno al array.
    Indice = Indice + 1
    
InicioControl:
    On Error GoTo errIS
    Load Sockets(Indice)
    ws_InstanciaSocket = Indice
    Exit Function
    
errIS:
    Indice = -1
    Resume Next
End Function

Private Sub ws_Desconecto(Socket As Winsock)
    On Error Resume Next
    If Socket.State <> sckClosed Then Socket.Close
End Sub

Private Function ws_CheckSocketsStates()
On Error GoTo errVS
Dim mRows As Integer, mIDSock As Integer, mRow As Integer
Dim mIndex As Integer

    With vsConexion
        mRows = .Rows - 1
        mRow = .FixedRows
        
        For mIndex = .FixedRows To mRows
            
            mIDSock = .Cell(flexcpText, mRow, 0)
            
            If wsSocket(mIDSock).State = sckError Or wsSocket(mIDSock).State = sckClosed Then
                ws_CierroConexion mIDSock
            Else
                mRow = mRow + 1
            End If
        Next
    End With
    Exit Function
    
errVS:
    loc_InsertoError "Método ValidateStates"
End Function

Private Sub wsSocket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    loc_InsertoError "wsSocket_Error (" & Index & "): " & Number & "- " & Description & "- Source: " & Source
    'reinicializamos
    ReInicializar
End Sub


Private Sub ReInicializar()
On Error Resume Next

    tmT.Enabled = False

    'Pelamos el array de sockets.
    Dim Socket As Winsock
    For Each Socket In wsSocket
        If Socket.Index <> 0 Then ws_Desconecto Socket
    Next
    
    vsConexion.Rows = 1
    
    'Prendemos el timer nuevamente.
    tmT.Enabled = True
    
End Sub
