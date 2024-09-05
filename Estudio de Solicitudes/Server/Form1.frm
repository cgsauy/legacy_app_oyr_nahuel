VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Servidor de Comunicación"
   ClientHeight    =   6645
   ClientLeft      =   3360
   ClientTop       =   2340
   ClientWidth     =   8025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   8025
   Begin VB.Timer tmT 
      Enabled         =   0   'False
      Left            =   540
      Top             =   4800
   End
   Begin VB.PictureBox picImagen 
      BorderStyle     =   0  'None
      Height          =   2355
      Index           =   2
      Left            =   5520
      ScaleHeight     =   2355
      ScaleWidth      =   2775
      TabIndex        =   13
      Top             =   780
      Width           =   2775
      Begin VB.PictureBox picIcono 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   2
         Left            =   120
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   14
         Top             =   60
         Width           =   480
      End
      Begin VSFlex6DAOCtl.vsFlexGrid vsLine 
         Height          =   1575
         Left            =   60
         TabIndex        =   15
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
      Begin VB.Label Label4 
         Caption         =   "Informacion"
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
         TabIndex        =   17
         Top             =   180
         Width           =   1935
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Height          =   45
         Left            =   0
         TabIndex        =   16
         Top             =   600
         Width           =   4455
      End
   End
   Begin VB.CommandButton bPrueba 
      Caption         =   "Prueba"
      Height          =   315
      Left            =   5340
      TabIndex        =   12
      Top             =   6300
      Width           =   1215
   End
   Begin VB.CommandButton bSalir 
      Caption         =   "&Salir"
      Height          =   315
      Left            =   6900
      TabIndex        =   11
      Top             =   6300
      Width           =   1035
   End
   Begin VB.PictureBox picImagen 
      BorderStyle     =   0  'None
      Height          =   2355
      Index           =   1
      Left            =   3660
      ScaleHeight     =   2355
      ScaleWidth      =   2775
      TabIndex        =   3
      Top             =   600
      Width           =   2775
      Begin VB.PictureBox picIcono 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   120
         Picture         =   "Form1.frx":08CA
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
         Caption         =   "Errores"
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
      Left            =   120
      ScaleHeight     =   3015
      ScaleWidth      =   4635
      TabIndex        =   1
      Top             =   600
      Width           =   4635
      Begin VB.PictureBox picIcono 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   0
         Left            =   120
         Picture         =   "Form1.frx":1194
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
         Caption         =   "Conexiones activas"
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
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2884
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Conexiones"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Errores"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Info"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock wsSocket 
      Index           =   0
      Left            =   4320
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aRow As Integer

Private Sub bPrueba_Click()
    'Dim i As Integer
    'If vsConexion.Rows = 1 Then Exit Sub
    'For i = 1 To vsConexion.Rows - 1
    '    If wsSocket(vsConexion.Cell(flexcpText, i, 0)).State = 7 Then wsSocket(vsConexion.Cell(flexcpText, i, 0)).SendData "LA REPARIO"
    '    DoEvents
    'Next i
    aRow = 1
    If aRow = 1 Then tmT.Interval = 1000
    tmT.Enabled = True
    
    
End Sub

Private Sub bSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    With vsConexion
        .Rows = 1
        .Cols = 4
        .FormatString = ">  ID  |<Dirección IP|>Puerto Remoto|>Puerto Local"
        .ColWidth(1) = 1700
    End With
    With vsError
        .Rows = 1
        .Cols = 2
        .FormatString = "Hora|Error"
        .ColWidth(0) = 780
    End With
    vsLine.LoadGrid "Buffer.txt", flexFileAll
    
    'Le asigno el puerto al socket
    wsSocket(0).LocalPort = 1251
    wsSocket(0).Listen
    '...........................................................
    tabServer.Left = 25
    tabServer.Top = 25
    picImagen(0).ZOrder 0
    Exit Sub
errLoad:
    MsgBox "Ocurrió un error al iniciar el formulario." _
        & vbCrLf & "Error: " & Trim(Err.Description), vbCritical, "ATENCIÓN"
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    tabServer.Width = Me.ScaleWidth - 50
    tabServer.Height = Me.ScaleHeight - (bSalir.Height + 120)
    With picImagen(0)
        .Left = tabServer.ClientLeft
        .Width = tabServer.ClientWidth
        .Height = tabServer.ClientHeight
        .Top = tabServer.ClientTop
    End With
    With picImagen(1)
        .Left = tabServer.ClientLeft
        .Width = tabServer.ClientWidth
        .Height = tabServer.ClientHeight
        .Top = tabServer.ClientTop
    End With
    With picImagen(2)
        .Left = tabServer.ClientLeft
        .Width = tabServer.ClientWidth
        .Height = tabServer.ClientHeight
        .Top = tabServer.ClientTop
    End With
    
    
    With vsConexion
        .Left = 30
        .Width = picImagen(0).Width - 120
        .Height = picImagen(0).Height - (vsConexion.Top + 50)
    End With
    labSeparador.Width = vsConexion.Width
    With vsError
        .Left = 30
        .Width = picImagen(0).Width - 120
        .Height = picImagen(0).Height - (vsError.Top + 50)
    End With
    
    With vsLine
        .Left = 30
        .Width = picImagen(0).Width - 120
        .Height = picImagen(0).Height - (vsError.Top + 50)
    End With
    
    labSeparador.Left = vsError.Left
    labSeparador1.Left = vsError.Left
    labSeparador1.Width = vsError.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Socket As Winsock
    For Each Socket In wsSocket
        pct_Desconecto Socket
    Next
End Sub

Private Sub tabServer_Click()
    picImagen(tabServer.SelectedItem.index - 1).ZOrder 0
End Sub

Private Sub tmT_Timer()
    Dim i As Integer
    If vsConexion.Rows = 1 Then Exit Sub
    vsLine.Cell(flexcpBackColor, aRow, 0, , vsLine.Cols - 1) = vbInactiveBorder
    vsLine.TopRow = aRow
    
    For i = 1 To vsConexion.Rows - 1
        If wsSocket(vsConexion.Cell(flexcpText, i, 0)).State = 7 Then wsSocket(vsConexion.Cell(flexcpText, i, 0)).SendData vsLine.Cell(flexcpText, aRow, 4)
        DoEvents
    Next i
    
    If aRow = vsLine.Rows Then
        tmT.Enabled = False
    Else
        On Error Resume Next
        tmT.Interval = Abs(DateDiff("s", vsLine.Cell(flexcpText, aRow, 0), vsLine.Cell(flexcpText, aRow + 1, 0))) * 1000
        If tmT.Interval = 0 Then tmT.Interval = 1000
        aRow = aRow + 1

        If aRow = vsLine.Rows Then tmT.Enabled = False Else tmT.Enabled = True
    End If
    
    
    
End Sub

Private Sub wsSocket_Close(index As Integer)
    CierroConexionSocket index
End Sub

Private Sub wsSocket_ConnectionRequest(index As Integer, ByVal requestID As Long)
Dim lngNuevoIndice As Long
    'Siempre escucho con el control cero y le asigno la conexión a un nuevo control.
    If index = 0 Then
        'Retorno el primer socket libre o asigno uno nuevo al array.
        lngNuevoIndice = pct_InstanciaSocket(wsSocket)
        If lngNuevoIndice > 0 Then
            wsSocket(lngNuevoIndice).LocalPort = 0      'Asigna el primer puerto disponible.
            wsSocket(lngNuevoIndice).Accept requestID
            InsertoConexionEnLista lngNuevoIndice
        Else
            InsertoError "Nueva conexión. " & Trim(Err.Description)
        End If
    End If
End Sub

Private Sub wsSocket_DataArrival(index As Integer, ByVal bytesTotal As Long)
Dim strDato As String

    If wsSocket(index).BytesReceived > 0 Then
        wsSocket(index).GetData strDato
        
        'Casos de ingreso de datos son Login y LogOff
        Select Case Trim(strDato)
            Case enuEventoPCT.Login & strSeparador
                    LoginOK index
            Case enuEventoPCT.LogOff & strSeparador
                    CierroConexionSocket index
        End Select
        
    End If
    
End Sub

Private Sub InsertoConexionEnLista(ByVal index As Long)
    With vsConexion
        .AddItem index
        .Cell(flexcpText, .Rows - 1, 1) = wsSocket(index).RemoteHostIP
        .Cell(flexcpText, .Rows - 1, 2) = wsSocket(index).RemotePort
        .Cell(flexcpText, .Rows - 1, 3) = wsSocket(index).LocalPort
    End With
End Sub

Private Sub RemuevoConexionEnLista(ByVal index As Long)
Dim intCont As Integer
On Error Resume Next
    For intCont = 1 To vsConexion.Rows - 1
        If Val(vsConexion.Cell(flexcpText, intCont, 0)) = index Then vsConexion.RemoveItem intCont: Exit For
    Next intCont
End Sub

Private Sub CierroConexionSocket(ByVal index As Long)
On Error Resume Next
    
    'Cierro conexión y elimino de la lista.
    RemuevoConexionEnLista index
    pct_Desconecto wsSocket(index)
    Unload wsSocket(index)
    
End Sub

Private Sub InsertoError(ByVal strMsg As String)
    With vsError
        .AddItem Time
        .Cell(flexcpText, 1, 1) = strMsg
    End With
End Sub

Private Sub LoginOK(ByVal index As Long)
    wsSocket(index).SendData enuEventoPCT.Login & strSeparador
End Sub
