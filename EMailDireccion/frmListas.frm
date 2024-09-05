VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACOMBO.OCX"
Begin VB.Form frmListas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listas de Distribución"
   ClientHeight    =   2940
   ClientLeft      =   4215
   ClientTop       =   4665
   ClientWidth     =   4020
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
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton bGrabar 
      Caption         =   "&Aceptar"
      Height          =   315
      Left            =   3000
      TabIndex        =   3
      Top             =   2580
      Width           =   975
   End
   Begin AACombo99.AACombo cLista 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   300
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      ListIndex       =   -1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
      Height          =   1815
      Left            =   60
      TabIndex        =   2
      Top             =   720
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   3201
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   4
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
   Begin VB.Label lDireccion 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Lista:"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   675
   End
End
Attribute VB_Name = "frmListas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public prmIDEMail As Long
Public prmDireccion As String

Private Sub AgregoLista()
    
    If cLista.ListIndex = -1 Then Exit Sub
    
    Dim I As Integer, aValor As Long, bHay As Boolean
    aValor = cLista.ItemData(cLista.ListIndex)
    bHay = False
    
    With vsLista
        For I = 1 To .Rows - 1
            If .Cell(flexcpData, I, 0) = aValor Then bHay = True
        Next
        
        If bHay Then Exit Sub

        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = Trim(cLista.Text)
        .Cell(flexcpData, .Rows - 1, 0) = aValor
        .Cell(flexcpText, .Rows - 1, 1) = "1"
        
        cLista.Text = ""
        cLista.SetFocus
    End With

End Sub

Private Sub bGrabar_Click()
    On Error GoTo errGrabar
    Screen.MousePointer = 11

    'Grabo los nuevos valores------------------------------------------------------------------
    Dim I As Integer, bHay As Boolean
    
    With vsLista
        bHay = False
        For I = 1 To .Rows - 1
            If .Cell(flexcpText, .Rows - 1, 1) = "1" Then bHay = True: Exit For
        Next I
        
        If bHay Then
            cons = "Select * From EMailLista Where EMLMail =" & prmIDEMail
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            For I = 1 To .Rows - 1
                If .Cell(flexcpText, I, 1) = "1" Then
                    rsAux.AddNew
                    rsAux!EMLMail = prmIDEMail
                    rsAux!EMLLista = .Cell(flexcpData, I, 0)
                    rsAux!EMLFAlta = Format(Now, "mm/dd/yyyy hh:mm:ss")
                    rsAux.Update
                End If
            Next
            rsAux.Close
        End If
    End With
    '----------------------------------------------------------------------------------------------
    
    Unload Me
    Screen.MousePointer = 0
    Exit Sub
    
errGrabar:
    clsGeneral.OcurrioError "Error al grabar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub cLista_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cLista.ListIndex <> -1 Then AgregoLista Else bGrabar.SetFocus
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    InicializoControles
    CargoLista
    
    lDireccion.Caption = Trim(prmDireccion)
    Screen.MousePointer = 0
    
End Sub

Private Sub CargoLista()

    On Error GoTo errCargar
    Screen.MousePointer = 11
    Dim aValor As Long
    vsLista.Rows = 1
    cons = "Select * From EMailLista, ListaDistribucion " & _
                " Where EMLMail =" & prmIDEMail & _
                " And EMLLista = LiDCodigo " & _
                " Order by LiDNombre"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        With vsLista
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Trim(rsAux!LiDNombre)
            aValor = rsAux!LiDCodigo: .Cell(flexcpData, .Rows - 1, 0) = aValor
            
            rsAux.MoveNext
        End With
    Loop
    rsAux.Close
    Screen.MousePointer = 0
    
    Exit Sub
errCargar:
    clsGeneral.OcurrioError "Error al cargar la lista.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub InicializoControles()

    On Error Resume Next
    cons = "Select LiDCodigo, LiDNombre from ListaDistribucion Where LiDHabilitado = 1 Order by LiDNombre"
    CargoCombo cons, cLista
    
    With vsLista
        .Cols = 1: .Rows = 1
        .FormatString = "<Lista de Distribución|Nuevo"
            
        .WordWrap = True
        .ColHidden(1) = True
        .ColWidth(0) = 1800
        
        .ExtendLastCol = True: .FixedCols = 0
    End With
      
End Sub

Private Sub vsLista_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsLista.Rows = 1 Then Exit Sub
    
    If KeyCode = vbKeyDelete Then
        Screen.MousePointer = 11
        
        If Trim(vsLista.Cell(flexcpText, vsLista.Row, 1)) = "" Then
            cons = "Select * From EMailLista " & _
                       " Where EMLMail =" & prmIDEMail & _
                       " And EMLLista = " & vsLista.Cell(flexcpData, vsLista.Row, 0)
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            If Not rsAux.EOF Then rsAux.Delete
            rsAux.Close
        End If
        
        vsLista.RemoveItem vsLista.Row
        
        Screen.MousePointer = 0
    End If
    Exit Sub

errEliminar:
    clsGeneral.OcurrioError "Error al eliminar la dirección de la lista.", Err.Description
    Screen.MousePointer = 0
End Sub

