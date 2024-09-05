VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmListaResult 
   BackColor       =   &H8000000C&
   Caption         =   "Proceso de Resolución"
   ClientHeight    =   4110
   ClientLeft      =   3420
   ClientTop       =   3165
   ClientWidth     =   5715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListaResult.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   5715
   Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
      Height          =   3195
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   5636
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
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
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
   Begin VB.Label lResult 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   60
      Width           =   3675
   End
End
Attribute VB_Name = "frmListaResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public prmColR As Collection
Public prmIdR As Long

Private Sub Form_Load()
    'Center form
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    lResult.Caption = "Resulución Automática NO pudo resolver la solicitud..."
    If prmIdR <> 0 Then
        cons = "Select * from CondicionResolucion Where ConCodigo = " & prmIdR
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            lResult.Caption = Trim(rsAux!ConNombre)
        End If
        rsAux.Close
    End If
    
    InicializoGrilla
    CargoLista
    Screen.MousePointer = 0
    
End Sub

Private Sub CargoLista()

Dim aTexto As String, aItem As String
    For I = 1 To prmColR.Count
        With vsLista
            aItem = prmColR(I)
            .AddItem ""
            
            aTexto = Mid(aItem, 1, InStr(aItem, "|") - 1)
            aItem = Mid(aItem, InStr(aItem, "|") + 1)
            .Cell(flexcpText, .Rows - 1, 1) = aTexto
            
            aTexto = Mid(aItem, 1, InStr(aItem, "|") - 1)
            aItem = Mid(aItem, InStr(aItem, "|") + 1)
            .Cell(flexcpText, .Rows - 1, 2) = aTexto
            
            .Cell(flexcpFontBold, .Rows - 1, 0) = True
            If Trim(aItem) = "" Then
                .Cell(flexcpText, .Rows - 1, 0) = "ERR"
            Else
                If Not CBool(aItem) Then
                    '.Cell(flexcpBackColor, .Rows - 1, 0) = vbRed
                    .Cell(flexcpText, .Rows - 1, 0) = " NO "
                Else
                    .Cell(flexcpText, .Rows - 1, 0) = " SI "
                End If
            End If
        End With
    Next
    
    Set prmColR = Nothing
    
End Sub

Private Sub InicializoGrilla()

    On Error Resume Next
    With vsLista
        .Cols = 1: .Rows = 1
        .FormatString = "|<Condición|<Expresión"
            
        .WordWrap = True
        .ColWidth(0) = 400: .ColWidth(1) = 1800: .ColWidth(2) = 2800
        
        .ExtendLastCol = True: .FixedCols = 0
    End With
      
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    lResult.Left = Me.ScaleLeft: lResult.Top = Me.ScaleTop + 50
    lResult.Width = Me.ScaleWidth
    
    With vsLista
        .Left = lResult.Left: .Width = lResult.Width
        .Top = lResult.Top + lResult.Height
        .Height = Me.ScaleHeight - .Top
    End With
    
End Sub
