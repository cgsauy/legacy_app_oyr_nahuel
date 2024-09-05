VERSION 5.00
Object = "{B443E3A5-0B4D-4B43-B11D-47B68DC130D7}#1.1#0"; "orArticulo.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmAviso 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Aviso Llegada"
   ClientHeight    =   3375
   ClientLeft      =   4125
   ClientTop       =   2640
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAviso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bCancel 
      Caption         =   "Cancelar"
      Height          =   350
      Left            =   3900
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton bGrabar 
      Caption         =   "Agregar"
      Height          =   350
      Left            =   60
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CheckBox cSimilares 
      Caption         =   "Similares"
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   780
      Width           =   1035
   End
   Begin prjFindArticulo.orArticulo tArticulo 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   720
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
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
      Height          =   1575
      Left            =   60
      TabIndex        =   7
      Top             =   1740
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2778
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
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
      BackColorFixed  =   -2147483645
      ForeColorFixed  =   -2147483634
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14737632
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   8421631
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   4
      GridLinesFixed  =   5
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
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
   Begin VB.Label lMail 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   4935
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   15
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   4995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Artículo"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   1035
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "?"
      Visible         =   0   'False
      Begin VB.Menu MnuHlp 
         Caption         =   "Ayuda"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmAviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bCancel_Click()
    Unload Me
End Sub

Private Sub bGrabar_Click()
    AccionGrabar
End Sub

Private Sub cSimilares_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bGrabar.SetFocus
End Sub

Private Sub Form_Load()

    InicializoForm
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EndMain
End Sub

Private Sub InicializoForm()
    
    On Error Resume Next
    LimpioFicha
    
    Set tArticulo.Connect = cBase
    tArticulo.DisplayCodigoArticulo = False
    
    mSQL = "Select (RTrim(EMDDireccion) + '@' + RTrim(EMSDireccion))  as Direccion " & _
                " From CGSA.dbo.EMailDireccion, CGSA.dbo.EMailServer" & _
                " Where EMDCodigo = " & prmIDMail & _
                " And EMDServidor = EMSCodigo"

    Set rsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        lMail.Caption = Trim(rsAux(0))
    End If
    rsAux.Close
    
    
    With vsLista
        .Rows = 1: .Cols = 1
        .FormatString = "^Incluído|<Artículos"
        .ColWidth(0) = 1100
        .WordWrap = False
        .MergeCells = flexMergeSpill
        .ExtendLastCol = True
        
        .Editable = True
        .RowHeight(0) = 280
        .SelectionMode = flexSelectionByRow
    End With
       
    fnc_CargoListas
    
End Sub

Private Function fnc_CargoListas()

    vsLista.Rows = vsLista.FixedRows
    
    Dim mValor As Long
    mSQL = "Select *, ArtNombre from CGSA.dbo.AvisoLlegada, CGSA.dbo.Articulo" & _
                " Where ALlEmail = " & prmIDMail & _
                " And ALlFechaNotificado IS NULL " & _
                " And ALlArticulo = ArtID"
    Set rsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        With vsLista
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Format(rsAux("ALlFechaIncluido"), "dd/mm hh:mm")
                
            .Cell(flexcpText, .Rows - 1, 1) = Trim(rsAux("ArtNombre"))
            mValor = rsAux("ALlArticulo"): .Cell(flexcpData, .Rows - 1, 1) = mValor
            If Not IsNull(rsAux("ALlArtSimilares")) Then .Cell(flexcpText, .Rows - 1, 1) = .Cell(flexcpText, .Rows - 1, 1) & " y Similares"
        End With
        rsAux.MoveNext
    Loop
    rsAux.Close

End Function


Private Sub LimpioFicha()
    tArticulo.Text = ""
    cSimilares.Value = vbUnchecked
End Sub

Private Function ValidoGrabar() As Boolean
On Error GoTo errValidar

    ValidoGrabar = False
        
    If tArticulo.prm_ArtID = 0 Then tArticulo.SetFocus: Exit Function
            
    If tArticulo.GetField("ArtHabilitado") = "S" Then
        
        If MsgBox("El artículo ingresado está habilitado para la venta." & vbCrLf & _
                        "¿Está seguro que quiere agregar el aviso a éste artículo?", vbQuestion + vbYesNo + vbDefaultButton2, "Artículo Habilitado") = vbNo Then
                        Exit Function
        End If
    End If
    
    
    ValidoGrabar = True
    
errValidar:
End Function

Private Sub AccionGrabar()
On Error GoTo errGrabar
   
    If Not ValidoGrabar Then Exit Sub
    prmIDArticulo = tArticulo.prm_ArtID
    'If MsgBox("¿Confirma agregar el mails a la lista de avisos?", vbQuestion + vbYesNo, "Agregar Aviso") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    
    Dim mSQL As String, rsAdd As rdoResultset
    
    mSQL = "Select * from CGSA.dbo.AvisoLlegada " & _
                " Where ALlEmail = " & prmIDMail & " And ALlArticulo = " & prmIDArticulo & _
                " And ALlFechaNotificado IS NULL"
                
    Set rsAdd = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    If rsAdd.EOF Then
        rsAdd.AddNew
        rsAdd!ALlEmail = prmIDMail
        rsAdd!ALlArticulo = prmIDArticulo
        rsAdd!ALlArtSimilares = IIf(cSimilares.Value = vbChecked, 1, Null)
        rsAdd!ALlFechaIncluido = Now
        rsAdd.Update
    End If
    rsAdd.Close

    Screen.MousePointer = 11
    'Unload Me
    fnc_CargoListas
    Screen.MousePointer = 0
    
    On Error Resume Next
    LimpioFicha
    tArticulo.SetFocus
    Exit Sub

errGrabar:
    clsGeneral.OcurrioError "Error al grabar los datos.", Err.Description
    Screen.MousePointer = 0: Exit Sub
End Sub

Private Sub tArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cSimilares.SetFocus
End Sub

Private Sub vsLista_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsLista.Rows = vsLista.FixedRows Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyDelete
        
            If MsgBox("Lista Aviso: " & vsLista.Cell(flexcpText, vsLista.Row, 1) & vbCrLf & vbCrLf & "¿Confirma eliminar el mail a la lista de avisos seleccionada?", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar Aviso") = vbNo Then Exit Sub
            
            mSQL = "Select * from CGSA.dbo.AvisoLlegada " & _
                        "Where ALlEmail = " & prmIDMail & " And ALlArticulo = " & vsLista.Cell(flexcpData, vsLista.Row, 1)
            Set rsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
            If Not rsAux.EOF Then rsAux.Delete
            rsAux.Close
            
            vsLista.RemoveItem vsLista.Row
            
    End Select
    
End Sub

