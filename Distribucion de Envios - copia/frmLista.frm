VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form frmLista 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin VSFlex6DAOCtl.vsFlexGrid vsGridArt 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1931
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
      BackColor       =   15794175
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483636
      ForeColorFixed  =   -2147483634
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16448250
      BackColorAlternate=   15791610
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483636
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   400
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   "Q|Artículo"
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
End
Attribute VB_Name = "frmLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub s_FillGrid(ByVal lEnvio As Long)
Dim rsFG As rdoResultset
    Screen.MousePointer = 11
    vsGridArt.Rows = 1
    Cons = "Select Sum(REvCantidad), ArtCodigo, rTrim(ArtNombre)" & _
                " From Envio, RenglonEnvio, Articulo" & _
                " Where ((EnvCodigo = " & lEnvio & " and EnvVaCon Is Null) or (abs(EnvVaCon) = " & lEnvio & "))" & _
                " And EnvCodigo = REvEnvio And REvArticulo = ArtID Group By ArtCodigo, ArtNombre"
    Set rsFG = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsFG.EOF
        With vsGridArt
            .AddItem rsFG(0)
            .Cell(flexcpText, .Rows - 1, 1) = Format(rsFG(1), "(000,000)") & " " & rsFG(2)
        End With
        rsFG.MoveNext
    Loop
    rsFG.Close
    Screen.MousePointer = 0
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_Load()
    Hook Me.hwnd
End Sub

Private Sub Form_Resize()
    vsGridArt.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    vsGridArt.Rows = 1
    Unhook
    frmDistribuirEnvio.vsGrid.SetFocus
End Sub

Private Sub vsGridArt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: Unload Me
    End Select
End Sub

