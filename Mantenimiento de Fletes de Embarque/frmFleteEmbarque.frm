VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{F93D243E-5C15-11D5-A90D-000021860458}#10.0#0"; "orFecha.ocx"
Begin VB.Form frmFleteEmbarque 
   Caption         =   "Mantenimiento de fletes"
   ClientHeight    =   5700
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   14010
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFleteEmbarque.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   14010
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar tooMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   14010
      _ExtentX        =   24712
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   11
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir del formulario"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "find"
            Object.ToolTipText     =   "Buscar fletes. [Ctrl+F]"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox txtTT 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3240
      MaxLength       =   3
      TabIndex        =   17
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtConjOrigen 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F4F1&
      Height          =   315
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox txtConjLinea 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F4F1&
      Height          =   315
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox tGastosMvd 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1140
      MaxLength       =   10
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox tDiasFree 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12120
      MaxLength       =   2
      TabIndex        =   13
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox tMemo 
      Height          =   525
      Left            =   780
      MaxLength       =   150
      MultiLine       =   -1  'True
      TabIndex        =   25
      Text            =   "frmFleteEmbarque.frx":08CA
      Top             =   1800
      Width           =   11895
   End
   Begin AACombo99.AACombo cLinea 
      Height          =   315
      Left            =   780
      TabIndex        =   7
      Top             =   960
      Width           =   1995
      _ExtentX        =   3519
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
   Begin orctFecha.orFecha tFAPartir 
      Height          =   285
      Left            =   4680
      TabIndex        =   19
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Object.Width           =   1215
      EnabledMes      =   -1  'True
      EnabledAño      =   -1  'True
      FechaFormato    =   "dd/mm/yyyy"
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsDato 
      Height          =   2775
      Left            =   60
      TabIndex        =   28
      Top             =   2400
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   4895
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
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
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   -1  'True
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
   End
   Begin VB.TextBox tImporte 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   9720
      TabIndex        =   11
      Top             =   960
      Width           =   1215
   End
   Begin AACombo99.AACombo cContenedor 
      Height          =   315
      Left            =   7020
      TabIndex        =   9
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      BackColor       =   12648447
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
   Begin AACombo99.AACombo cAgencia 
      Height          =   315
      Left            =   9720
      TabIndex        =   5
      Top             =   600
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   556
      BackColor       =   12648447
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
   Begin AACombo99.AACombo cDestino 
      Height          =   315
      Left            =   7020
      TabIndex        =   3
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      BackColor       =   12648447
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
   Begin AACombo99.AACombo cOrigen 
      Height          =   315
      Left            =   780
      TabIndex        =   1
      Top             =   600
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   556
      BackColor       =   12648447
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
   Begin ComctlLib.StatusBar staMsg 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   27
      Top             =   5445
      Width           =   14010
      _ExtentX        =   24712
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin orctFecha.orFecha tEmbarco 
      Height          =   285
      Left            =   9720
      TabIndex        =   23
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Object.Width           =   1215
      EnabledMes      =   -1  'True
      EnabledAño      =   -1  'True
      FechaFormato    =   "dd/mm/yyyy"
   End
   Begin orctFecha.orFecha tFHasta 
      Height          =   285
      Left            =   7080
      TabIndex        =   21
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Object.Width           =   1215
      EnabledMes      =   -1  'True
      EnabledAño      =   -1  'True
      FechaFormato    =   "dd/mm/yyyy"
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "TT:"
      Height          =   195
      Left            =   2880
      TabIndex        =   16
      ToolTipText     =   "Transit time"
      Top             =   1440
      Width           =   315
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Gastos Mvd:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Días libre:"
      Height          =   195
      Left            =   11280
      TabIndex        =   12
      Top             =   960
      Width           =   915
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Válido hasta:"
      Height          =   195
      Left            =   6000
      TabIndex        =   20
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "&Memo:"
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   1800
      Width           =   675
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "&Embarcó:"
      Height          =   195
      Left            =   8880
      TabIndex        =   22
      Top             =   1440
      Width           =   795
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "&Línea:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   675
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "P&recio:"
      Height          =   195
      Left            =   8880
      TabIndex        =   10
      Top             =   960
      Width           =   675
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Contenedor:"
      Height          =   195
      Left            =   6000
      TabIndex        =   8
      Top             =   960
      Width           =   915
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "A &partir:"
      Height          =   195
      Left            =   3960
      TabIndex        =   18
      Top             =   1440
      Width           =   675
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Agencia:"
      Height          =   195
      Left            =   8880
      TabIndex        =   4
      Top             =   600
      Width           =   675
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Destino:"
      Height          =   195
      Left            =   6300
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Origen:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   675
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5100
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFleteEmbarque.frx":092F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFleteEmbarque.frx":0A41
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFleteEmbarque.frx":0B53
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFleteEmbarque.frx":0C65
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFleteEmbarque.frx":0D77
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFleteEmbarque.frx":1091
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFleteEmbarque.frx":13AB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuOpLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGrabar 
         Caption         =   "&Grabar"
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "&Cancelar"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuOpLinea2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpFind 
         Caption         =   "Buscar Fletes"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuSalForm 
         Caption         =   "Del formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmFleteEmbarque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bEdicion As Boolean
Dim idFlete As Long
Dim aLinea() As String, aAgencia() As String, aCiudad() As String, aCont() As String

Dim listaOrigen As clsListaCodigoNombre
Dim listaLinea As clsListaCodigoNombre

Public prmPrm As String

Private Sub ProcesoTextoLista(ByVal textB As TextBox, ByVal Lista As clsListaCodigoNombre)
On Error Resume Next
    If textB.SelLength > 0 And Not listaOrigen Is Nothing Then
        Dim sSel As String
        sSel = textB.SelText
        If sSel = Trim(textB.Text) Then
            textB.Text = ""
            Set Lista.Lista = New Collection
        End If
        Dim iSel As Integer
        iSel = Lista.BuscarTexto(sSel)
        If iSel > 0 Then
            Lista.Lista.Remove iSel
            textB.Text = Lista.ListaElementosAsignado()
        End If
    End If

End Sub

Private Function fnc_GetIDArray(ByVal iValor As Byte, ByVal sTxt As String)
Dim aAux() As String
Dim iQ As Integer
    
    Select Case iValor
        Case 0
            fnc_GetIDArray = fnc_GetID(aCont, sTxt)
        Case 3
            fnc_GetIDArray = fnc_GetID(aAgencia, sTxt)
        Case 4, 5
            fnc_GetIDArray = fnc_GetID(aCiudad, sTxt)
        Case 7
            fnc_GetIDArray = fnc_GetID(aLinea, sTxt)
    End Select
    
End Function

Private Function fnc_GetID(ByVal aAux As Variant, ByVal sTxt As String) As Long
Dim iQ As Integer
    For iQ = 0 To UBound(aAux)
        If InStr(1, aAux(iQ), "Ã@¥" & sTxt, vbTextCompare) > 0 Then
            fnc_GetID = Mid(aAux(iQ), 1, InStr(1, aAux(iQ), "Ã@¥", vbTextCompare) - 1)
            Exit Function
        End If
    Next
End Function

Private Sub loc_SaveByGrid(ByVal iCol As Integer, ByVal idFlete As Long)
On Error GoTo errSC
    Cons = "Select * From FleteEmbarque Where FEmID = " & idFlete
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        MsgBox "El código que está editando fue eliminado, refresque la consulta.", vbExclamation, "Atención"
    Else
        RsAux.Edit
        Select Case iCol
            Case 0:
                RsAux!FEmContenedor = fnc_GetIDArray(iCol, vsDato.EditText)
            Case 1
                If IsDate(vsDato.EditText) Then RsAux!FEmFAPartir = Format(vsDato.EditText, "mm/dd/yyyy") Else RsAux!FEmFAPartir = Null
            Case 2
                If IsDate(vsDato.EditText) Then RsAux!FEmFHasta = Format(vsDato.EditText, "mm/dd/yyyy") Else RsAux!FEmFHasta = Null
            Case 3
                RsAux!FEmAgencia = fnc_GetIDArray(iCol, vsDato.EditText)
            Case 4
                RsAux!FEmOrigen = fnc_GetIDArray(iCol, vsDato.EditText)
            Case 5
                RsAux!FEmDestino = fnc_GetIDArray(iCol, vsDato.EditText)
            Case 6
                RsAux!FEmImporte = Format(vsDato.EditText, "###0.00")
                vsDato.EditText = Format(vsDato.EditText, "#,##0.00")
            Case 7
                RsAux!FEmLinea = fnc_GetIDArray(iCol, vsDato.EditText)
            Case 8
                If Trim(vsDato.EditText) = "" Then RsAux!FEmComentario = Null Else RsAux!FEmComentario = Trim(vsDato.EditText)
            Case 9
                RsAux("FEmDiasLibre") = IIf(vsDato.EditText = "", Null, vsDato.EditText)
            Case 10
                RsAux("FEmGastosMvd") = IIf(vsDato.EditText = "", Null, vsDato.EditText)
                If vsDato.EditText <> "" Then vsDato.EditText = Format(vsDato.EditText, "#,##0.00")
            Case 11
                RsAux("FEmDiasTransito") = IIf(vsDato.EditText = "", Null, vsDato.EditText)
        End Select
        RsAux.Update
    End If
    RsAux.Close
    Exit Sub
errSC:
    MsgBox "Error al grabar la edición: " & Err.Description, vbCritical, "Atención"
End Sub

Public Function FillComboGetGridCombo(Consulta As String, Combo As Control, ByRef sConIDs As String) As String
Dim RsAuxiliar As rdoResultset
On Error GoTo ErrCC
    
    Screen.MousePointer = 11
    sConIDs = ""
    Combo.Clear
    Set RsAuxiliar = cBase.OpenResultset(Consulta, rdOpenDynamic, rdConcurValues)
    Do While Not RsAuxiliar.EOF
        Combo.AddItem Trim(RsAuxiliar(1))
        Combo.ItemData(Combo.NewIndex) = RsAuxiliar(0)
        FillComboGetGridCombo = FillComboGetGridCombo & IIf(FillComboGetGridCombo = "", "", "|") & Trim(RsAuxiliar(1))
        sConIDs = sConIDs & IIf(sConIDs = "", "", "|") & Trim(RsAuxiliar(0)) & "Ã@¥" & Trim(RsAuxiliar(1))
        RsAuxiliar.MoveNext
    Loop
    RsAuxiliar.Close
    Screen.MousePointer = 0
    Exit Function
    
ErrCC:
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al cargar el combo: " & Trim(Combo.Name) & "." & vbCrLf & Err.Description, vbCritical, "ERROR"
End Function

Private Sub cAgencia_GotFocus()
    With cAgencia
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub cAgencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cLinea.SetFocus
End Sub

Private Sub cContenedor_GotFocus()
    With cContenedor
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub cContenedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tImporte.SetFocus
End Sub

Private Sub cDestino_GotFocus()
    With cDestino
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub cDestino_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cAgencia.SetFocus
End Sub

Private Sub cLinea_GotFocus()
On Error Resume Next
    With cLinea
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub cLinea_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If idFlete = 0 And cLinea.ListIndex > -1 And bEdicion Then
            'Agrego en el txt y dejo que me ingrese más.
            If listaLinea.BuscarElemento(cLinea.ItemData(cLinea.ListIndex)) Is Nothing Then
                Dim oCN As New clsCodigoNombre
                oCN.Codigo = cLinea.ItemData(cLinea.ListIndex)
                oCN.Nombre = cLinea.Text
                listaLinea.Lista.Add oCN
                txtConjLinea.Text = listaLinea.ListaElementosAsignado
                cLinea.Text = ""
            End If
        Else
            cContenedor.SetFocus
        End If
    End If
End Sub

Private Sub cOrigen_GotFocus()
    With cOrigen
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub cOrigen_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If idFlete = 0 And cOrigen.ListIndex > -1 And bEdicion Then
            'Agrego en el txt y dejo que me ingrese más.
            If listaOrigen.BuscarElemento(cOrigen.ItemData(cOrigen.ListIndex)) Is Nothing Then
                Dim oCN As New clsCodigoNombre
                oCN.Codigo = cOrigen.ItemData(cOrigen.ListIndex)
                oCN.Nombre = cOrigen.Text
                listaOrigen.Lista.Add oCN
                txtConjOrigen.Text = listaOrigen.ListaElementosAsignado
                cOrigen.Text = ""
            End If
        Else
            cDestino.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    ObtengoSeteoForm Me, 500, 500
    Me.Height = 6360
    
    Botones True, False, False, False, False, tooMenu, Me
    LimpioDatos
    bEdicion = False
    tMemo.Locked = Not bEdicion
    
    With vsDato
        .Rows = 1: .Cols = 1: .ExtendLastCol = True
        .FormatString = "<Conten.|A Partir|Valida Hasta|Agencia|Origen|Destino|>Importe|Línea|<Memo|>  Dias|>  Gastos Mvd|TT|"
        .ColWidth(1) = 950
        .ColDataType(1) = flexDTDate
        .ColDataType(2) = flexDTDate
        .ColDataType(6) = flexDTCurrency
        .ColWidth(3) = 1600
        .ColWidth(4) = 1300
        .ColWidth(5) = 1300
        .ColWidth(6) = 950
        .ColWidth(7) = 1000
        .ColWidth(8) = 4000
        .ColWidth(11) = 1000
        .ColWidth(12) = 10
    End With
    Dim sAux As String, sIDs As String
    
    'Cargo los Contenedores.-------------
    Cons = "Select ConCodigo, ConAbreviacion from Contenedor" _
        & " Order by ConNombre"
    sAux = FillComboGetGridCombo(Cons, cContenedor, sIDs)
    aCont = Split(sIDs, "|")
    vsDato.ColComboList(0) = sAux
    '----------------------------------------------
    
    sIDs = ""
    'Cargo Ciudad Origen y Ciudad Destino.
    Cons = "Select CiuCodigo, CiuNombre from Ciudad" _
        & " Order by CiuNombre"
    sAux = FillComboGetGridCombo(Cons, cOrigen, sIDs)
    sAux = FillComboGetGridCombo(Cons, cDestino, sIDs)
    vsDato.ColComboList(4) = sAux
    aCiudad = Split(sIDs, "|")
    vsDato.ColComboList(5) = sAux
    '----------------------------------------------
    
    'Cargo las Agencias de Transportes.--------------------
    Cons = "Select ATrCodigo, ATrNombre From AgenciaTransporte Order by ATrNombre"
    sAux = FillComboGetGridCombo(Cons, cAgencia, sIDs)
    vsDato.ColComboList(3) = sAux
    aAgencia = Split(sIDs, "|")
    '----------------------------------------------
    Cons = "Select Codigo, Texto From CodigoTexto Where Tipo = 67 Order by Texto"
    sAux = FillComboGetGridCombo(Cons, cLinea, sIDs)
    vsDato.ColComboList(7) = sAux
    aLinea = Split(sIDs, "|")
    '----------------------------------------------
    
    On Error Resume Next
    Dim aPrm() As String
    
    If Me.prmPrm <> "" Then
        aPrm = Split(Me.prmPrm, ",")
        Dim iQ As Integer
        For iQ = 0 To UBound(aPrm)
            If InStr(1, aPrm(iQ), ":") > 0 Then
                Select Case LCase(Mid(Trim(aPrm(iQ)), 1, 4))
                    Case "ori:": BuscoCodigoEnCombo cOrigen, Mid(Trim(aPrm(iQ)), 5)
                    Case "age:": BuscoCodigoEnCombo cAgencia, Mid(Trim(aPrm(iQ)), 5)
                    Case "lin:": BuscoCodigoEnCombo cLinea, Mid(Trim(aPrm(iQ)), 5)
                    Case "con:": BuscoCodigoEnCombo cContenedor, Mid(Trim(aPrm(iQ)), 5)
                End Select
            End If
        Next
    End If
    Screen.MousePointer = 0
    Exit Sub
errLoad:
    clsGeneral.OcurrioError "Ocurrió un error al iniciar el formulario.", Trim(Err.Description)
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    With vsDato
        .Width = Me.ScaleWidth - (.Left * 2)
        .Height = Me.ScaleHeight - .Top - staMsg.Height - 60
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    CierroConexion
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    GuardoSeteoForm Me
End Sub

Private Sub Label1_Click()
On Error Resume Next
    cOrigen.SetFocus
End Sub

Private Sub Label10_Click()
On Error Resume Next
    With tFHasta
        .SelStart = 0: .SelLength = Len(.FechaText): .SetFocus
    End With
End Sub

Private Sub Label2_Click()
On Error Resume Next
    cDestino.SetFocus
End Sub

Private Sub Label3_Click()
On Error Resume Next
    cAgencia.SetFocus
End Sub

Private Sub Label4_Click()
On Error Resume Next
    With tFAPartir
        .SelStart = 0: .SelLength = Len(.FechaText): .SetFocus
    End With
End Sub

Private Sub Label5_Click()
On Error Resume Next
    With cContenedor
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label6_Click()
On Error Resume Next
    With tImporte
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label7_Click()
On Error Resume Next
    With cLinea
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label9_Click()
    Foco tMemo
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuEliminar_Click()
    AccionEliminar
End Sub

Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuModificar_Click()
    AccionModificar
End Sub

Private Sub MnuNuevo_Click()
    AccionNuevo
End Sub

Private Sub MnuOpFind_Click()
    AccionBuscar
End Sub

Private Sub MnuSalForm_Click()
    Unload Me
End Sub

Private Sub tDiasFree_GotFocus()
    With tDiasFree
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tDiasFree_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then tGastosMvd.SetFocus
End Sub

Private Sub tDiasFree_LostFocus()
    tDiasFree.SelLength = 0
End Sub

Private Sub tDiasFree_Validate(Cancel As Boolean)
On Error Resume Next
    If Not IsNumeric(tDiasFree.Text) Then
        tDiasFree.Text = ""
    End If
End Sub

Private Sub tEmbarco_GotFocus()
    With tEmbarco
        .SelStart = 0: .SelLength = Len(.FechaText)
    End With
End Sub

Private Sub tEmbarco_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionBuscar
End Sub

Private Sub tFAPartir_GotFocus()
    With tFAPartir
        .SelStart = 0: .SelLength = Len(.FechaText)
    End With
End Sub

Private Sub tFAPartir_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tFHasta.SetFocus
End Sub

Private Sub tFHasta_GotFocus()
    With tFHasta
        .SelStart = 0: .SelLength = Len(.FechaText)
    End With
End Sub

Private Sub tFHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If bEdicion Then
            tMemo.SetFocus
        Else
            tEmbarco.SetFocus
        End If
    End If
End Sub

Private Sub tGastosMvd_GotFocus()
    With tGastosMvd
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tGastosMvd_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then txtTT.SetFocus
End Sub

Private Sub tGastosMvd_LostFocus()
    tGastosMvd.SelLength = 0
    On Error Resume Next
    If IsNumeric(tGastosMvd.Text) Then
        tGastosMvd.Text = Format(tGastosMvd.Text, "#,##0.00")
    Else
        tGastosMvd.Text = ""
    End If
End Sub

Private Sub tImporte_GotFocus()
    With tImporte
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tImporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tDiasFree.SetFocus
End Sub

Private Sub tMemo_GotFocus()
    With tMemo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tMemo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And MnuGrabar.Enabled Then
        If tMemo.SelLength > 0 Then tMemo.SelLength = 0
        KeyAscii = 0
        tMemo.Text = Replace(tMemo.Text, vbCrLf, "")
        AccionGrabar
    End If
End Sub

Private Sub tooMenu_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "salir": Unload Me
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        Case "find": AccionBuscar
    End Select
End Sub

Private Sub AccionNuevo()
    Botones False, False, False, True, True, tooMenu, Me
    vsDato.Enabled = False
    tEmbarco.Enabled = False
    LimpioDatos
    idFlete = 0
    bEdicion = True
    tMemo.Locked = Not bEdicion
    Botones False, False, False, True, True, tooMenu, Me
    cOrigen.SetFocus
End Sub

Private Sub AccionModificar()
    vsDato.Enabled = False
    LimpioDatos
    idFlete = vsDato.Cell(flexcpData, vsDato.Row, 0)
    bEdicion = True
    Cons = "Select * From FleteEmbarque Where FEmID = " & idFlete
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        idFlete = 0
        RsAux.Close
        MsgBox "Otra terminal pudo eliminar el flete seleccionado.", vbExclamation, "ATENCIÓN"
        Exit Sub
    Else
        tEmbarco.Enabled = False
        tMemo.Locked = Not bEdicion
        BuscoCodigoEnCombo cOrigen, RsAux!FEmOrigen
        BuscoCodigoEnCombo cDestino, RsAux!FEmDestino
        BuscoCodigoEnCombo cAgencia, RsAux!FEmAgencia
        BuscoCodigoEnCombo cContenedor, RsAux!FEmContenedor
        If Not IsNull(RsAux!FEmFAPartir) Then tFAPartir.FechaValor = Format(RsAux!FEmFAPartir, "dd/mm/yyyy") Else tFAPartir.FechaValor = ""
        If Not IsNull(RsAux!FEmFHasta) Then tFHasta.FechaValor = Format(RsAux!FEmFHasta, "dd/mm/yyyy") Else tFHasta.FechaValor = ""
        tImporte.Text = Format(RsAux!FEmImporte, "#,#00.00")
        If Not IsNull(RsAux!FEmLinea) Then BuscoCodigoEnCombo cLinea, RsAux!FEmLinea
        If Not IsNull(RsAux!FEmComentario) Then tMemo.Text = Trim(RsAux!FEmComentario) Else tMemo.Text = ""
        If Not IsNull(RsAux("FEmDiasLibre")) Then tDiasFree.Text = RsAux("FEmDiasLibre") Else tDiasFree.Text = ""
        If Not IsNull(RsAux("FEmGastosMvd")) Then tGastosMvd.Text = Format(RsAux("FEmGastosMvd"), "#,##0.00") Else tGastosMvd.Text = ""
        If Not IsNull(RsAux("FEmDiasTransito")) Then txtTT.Text = RsAux("FEmDiasTransito") Else txtTT.Text = ""
        RsAux.Close
    End If
    Botones False, False, False, True, True, tooMenu, Me
    cOrigen.SetFocus
End Sub

Private Sub AccionEliminar()

    If MsgBox("¿Confirma eliminar el flete seleccionado?", vbQuestion + vbYesNo, "Eliminar") = vbYes Then
        Screen.MousePointer = 11
        On Error GoTo errBT
        cBase.CommitTrans
        On Error GoTo errRB
        Cons = "Select * From FleteEmbarque Where FEmID = " & vsDato.Cell(flexcpData, vsDato.Row, 0)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.Delete
        RsAux.Close
        cBase.CommitTrans
        vsDato.Rows = 1
        Screen.MousePointer = 0
        AccionCancelar
    End If
    Exit Sub

errBT:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar la transacción.", Trim(Err.Description)
    Screen.MousePointer = 0
    Exit Sub

RollB:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrió un error al almacenar la información.", Trim(Err.Description)
    Screen.MousePointer = 0

errRB:
    Resume RollB

End Sub

Private Sub AccionGrabar()
    If MsgBox("¿Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
        If ValidoDatos Then GraboEnBD
    End If
End Sub

Private Sub AccionCancelar()
    bEdicion = False
    vsDato.Enabled = True
    tEmbarco.Enabled = True
    tMemo.Locked = Not bEdicion
    If cOrigen.ListIndex = -1 And cDestino.ListIndex = -1 And cAgencia.ListIndex = -1 Then
        Botones True, False, False, False, False, tooMenu, Me
    Else
        AccionBuscar 0
    End If
    LimpioDatos
    cOrigen.SetFocus
End Sub

Private Sub GraboEnBD()
Dim lpos As Long, lCol As Long

    Screen.MousePointer = 11
    On Error GoTo errBT
    '-----------------------------------------------------
    cBase.BeginTrans
    On Error GoTo errRB
    Cons = "Select * From FleteEmbarque Where FEmID = " & idFlete
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If idFlete > 0 Then
        RsAux.Edit
        RsAux!FEmOrigen = cOrigen.ItemData(cOrigen.ListIndex)
        RsAux!FEmDestino = cDestino.ItemData(cDestino.ListIndex)
        RsAux!FEmAgencia = cAgencia.ItemData(cAgencia.ListIndex)
        If IsDate(tFAPartir.FechaText) Then RsAux!FEmFAPartir = Format(tFAPartir.FechaValor, "mm/dd/yyyy") Else RsAux!FEmFAPartir = Null
        If IsDate(tFHasta.FechaText) Then RsAux!FEmFHasta = Format(tFHasta.FechaValor, "mm/dd/yyyy") Else RsAux!FEmFHasta = Null
        RsAux!FEmContenedor = cContenedor.ItemData(cContenedor.ListIndex)
        RsAux!FEmImporte = Format(tImporte.Text, "###0.00")
        If cLinea.ListIndex > -1 Then RsAux!FEmLinea = cLinea.ItemData(cLinea.ListIndex) Else RsAux!FEmLinea = Null
        If Trim(tMemo.Text) = "" Then RsAux!FEmComentario = Null Else RsAux!FEmComentario = Trim(tMemo.Text)
        RsAux("FEmDiasLibre") = IIf(tDiasFree.Text = "", Null, tDiasFree.Text)
        RsAux("FEmGastosMvd") = IIf(tGastosMvd.Text = "", Null, tGastosMvd.Text)
        RsAux("FEmDiasTransito") = IIf(txtTT.Text = "", Null, txtTT.Text)
        RsAux.Update
        RsAux.Close
    Else
        Dim oOrigen As clsCodigoNombre
        Dim oLinea As clsCodigoNombre
        
        For Each oOrigen In listaOrigen.Lista
            
            If listaLinea.Lista.Count = 0 Then
                RsAux.AddNew
                RsAux!FEmOrigen = oOrigen.Codigo
                RsAux!FEmDestino = cDestino.ItemData(cDestino.ListIndex)
                RsAux!FEmAgencia = cAgencia.ItemData(cAgencia.ListIndex)
                If IsDate(tFAPartir.FechaText) Then RsAux!FEmFAPartir = Format(tFAPartir.FechaValor, "mm/dd/yyyy") Else RsAux!FEmFAPartir = Null
                If IsDate(tFHasta.FechaText) Then RsAux!FEmFHasta = Format(tFHasta.FechaValor, "mm/dd/yyyy") Else RsAux!FEmFHasta = Null
                RsAux!FEmContenedor = cContenedor.ItemData(cContenedor.ListIndex)
                RsAux!FEmImporte = Format(tImporte.Text, "###0.00")
                RsAux!FEmLinea = Null
                If Trim(tMemo.Text) = "" Then RsAux!FEmComentario = Null Else RsAux!FEmComentario = Trim(tMemo.Text)
                RsAux("FEmDiasLibre") = IIf(tDiasFree.Text = "", Null, tDiasFree.Text)
                RsAux("FEmGastosMvd") = IIf(tGastosMvd.Text = "", Null, tGastosMvd.Text)
                RsAux("FEmDiasTransito") = IIf(txtTT.Text = "", Null, txtTT.Text)
                RsAux.Update
            Else
                For Each oLinea In listaLinea.Lista
                    RsAux.AddNew
                    RsAux!FEmOrigen = oOrigen.Codigo
                    RsAux!FEmDestino = cDestino.ItemData(cDestino.ListIndex)
                    RsAux!FEmAgencia = cAgencia.ItemData(cAgencia.ListIndex)
                    If IsDate(tFAPartir.FechaText) Then RsAux!FEmFAPartir = Format(tFAPartir.FechaValor, "mm/dd/yyyy") Else RsAux!FEmFAPartir = Null
                    If IsDate(tFHasta.FechaText) Then RsAux!FEmFHasta = Format(tFHasta.FechaValor, "mm/dd/yyyy") Else RsAux!FEmFHasta = Null
                    RsAux!FEmContenedor = cContenedor.ItemData(cContenedor.ListIndex)
                    RsAux!FEmImporte = Format(tImporte.Text, "###0.00")
                    RsAux!FEmLinea = oLinea.Codigo
                    If Trim(tMemo.Text) = "" Then RsAux!FEmComentario = Null Else RsAux!FEmComentario = Trim(tMemo.Text)
                    RsAux("FEmDiasLibre") = IIf(tDiasFree.Text = "", Null, tDiasFree.Text)
                    RsAux("FEmGastosMvd") = IIf(tGastosMvd.Text = "", Null, tGastosMvd.Text)
                    RsAux("FEmDiasTransito") = IIf(txtTT.Text = "", Null, txtTT.Text)
                    RsAux.Update
                Next
            End If
        Next
    End If
    cBase.CommitTrans
    '-----------------------------------------------------
    
    If idFlete = 0 Then
    
        Dim sLineas As String
        sLineas = listaLinea.ListaCodigos()
    
        On Error GoTo errQ
        Cons = "Select FEmID, FEmFAPartir, FEmFHasta, FEmImporte, CiuOrigen.CiuNombre as Origen, CiuDestino.CiuNombre as Destino, ATrNombre, ConAbreviacion, Texto, FEmComentario, FEmDiasLibre, FEmGastosMvd, FEmDiasTransito " _
                & " From FleteEmbarque " _
                    & " Left Outer Join CodigoTexto On Codigo = FEmLinea " _
                & ", Ciudad CiuOrigen, Ciudad CiuDestino, AgenciaTransporte, Contenedor" _
                & " Where FEmOrigen IN (" & listaOrigen.ListaCodigos & ")" _
                & " And FEmDestino = " & cDestino.ItemData(cDestino.ListIndex) _
                & " And FEmAgencia = " & cAgencia.ItemData(cAgencia.ListIndex) _
                & " And FEmContenedor = " & cContenedor.ItemData(cContenedor.ListIndex) _
                & " And FEmImporte = " & CCur(tImporte.Text) & " And FEmDiasLibre " & IIf(tDiasFree.Text <> "", " = " & tDiasFree.Text, " Is Null")

        If tFAPartir.FechaText <> "" Then Cons = Cons & " And FEmFAPartir >= '" & Format(tFAPartir.FechaText, "yyyy/mm/dd") & "'" _

        If sLineas <> "" Then Cons = Cons & " And FEmLinea IN(" & sLineas & ")"
            
        Cons = Cons & " And FEmOrigen = CiuOrigen.CiuCodigo And FEmDestino = CiuDestino.CiuCodigo " _
                & " And FEmAgencia = ATrCodigo And FEmContenedor = ConCodigo " _
                & " Order by CiuOrigen.CiuNombre"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            With vsDato
                .AddItem ""
                lpos = RsAux!FEmID: .Cell(flexcpData, .Rows - 1, 0) = lpos
                .Cell(flexcpText, .Rows - 1, 0) = Trim(RsAux!ConAbreviacion)
                If Not IsNull(RsAux!FEmFAPartir) Then .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!FEmFAPartir, "dd/mm/yyyy")
                If Not IsNull(RsAux!FEmFHasta) Then .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!FEmFHasta, "dd/mm/yyyy")
                .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!ATrNombre)
                .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!origen)
                .Cell(flexcpText, .Rows - 1, 5) = Trim(RsAux!Destino)
                .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!FEmImporte, "#,##0.00")
                If Not IsNull(RsAux!Texto) Then .Cell(flexcpText, .Rows - 1, 7) = Trim(RsAux!Texto)
                If Not IsNull(RsAux!FEmComentario) Then .Cell(flexcpText, .Rows - 1, 8) = Trim(RsAux!FEmComentario)
                If Not IsNull(RsAux("FEmDiasLibre")) Then .Cell(flexcpText, .Rows - 1, 9) = Trim(RsAux!FEmDiasLibre)
                If Not IsNull(RsAux("FEmGastosMvd")) Then .Cell(flexcpText, .Rows - 1, 10) = Format(RsAux!FEmGastosMvd, "#,##0.00")
                If Not IsNull(RsAux("FEmDiasTransito")) Then .Cell(flexcpText, .Rows - 1, 11) = RsAux!FEmDiasTransito
            End With
            RsAux.MoveNext
        Loop
        RsAux.Close
        'Me quedo con los datos que ingreso hasta el contenedor.
        cContenedor.Text = "": tImporte.Text = ""
        cContenedor.SetFocus
    Else
        On Error Resume Next
        bEdicion = False
        vsDato.Enabled = True
        tEmbarco.Enabled = True
        tMemo.Locked = Not bEdicion
        AccionBuscar 0, MantengoGrilla
    End If
    Screen.MousePointer = 0
    Exit Sub
errQ:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al restaurar la ficha.", Err.Description
    Exit Sub
    
errBT:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar la transacción.", Trim(Err.Description)
    Screen.MousePointer = 0
    Exit Sub

RollB:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrió un error al almacenar la información.", Trim(Err.Description)
    Screen.MousePointer = 0
    Exit Sub

errRB:
    Resume RollB
    Exit Sub
    
End Sub

Private Sub LimpioDatos()
    cOrigen.Text = ""
    cDestino.Text = ""
    cAgencia.Text = ""
    tImporte.Text = ""
    tFAPartir.FechaValor = ""
    tFHasta.FechaValor = ""
    tEmbarco.FechaValor = ""
    cContenedor.Text = ""
    cLinea.Text = ""
    tMemo.Text = ""
    tDiasFree.Text = ""
    tGastosMvd.Text = ""
    txtConjOrigen.Text = ""
    txtConjLinea.Text = ""
    
    Set listaOrigen = Nothing
    Set listaLinea = Nothing
    
    Set listaOrigen = New clsListaCodigoNombre
    Set listaLinea = New clsListaCodigoNombre
    
End Sub

Private Sub AccionBuscar(Optional Flete As Long = 0, Optional sIn As String)
On Error GoTo errAB
Dim lValor As Long, lInserto As Long
    
    If bEdicion Then Exit Sub
    
    If IsDate(tFAPartir.FechaText) Then tEmbarco.FechaValor = ""
    
    If IsDate(tEmbarco.FechaText) Then
        If cContenedor.ListIndex = -1 Then
            MsgBox "Debe seleccionar un Contenedor para ver la condición de Embarcó.", vbInformation, "ATENCIÓN"
            cContenedor.SetFocus
            Exit Sub
        End If
    End If
    vsDato.Rows = 1
    Botones True, False, False, False, False, tooMenu, Me
    If sIn <> "" Then
        Cons = "Select FEmID, FEmFAPartir, FEMFHasta, FEmImporte, CiuOrigen.CiuNombre as Origen, CiuDestino.CiuNombre as Destino, ATrNombre, ConAbreviacion, Texto, FEmComentario, FEmDiasLibre, FEmGastosMvd, FEmDiasTransito " _
                & " From FleteEmbarque " _
                    & " Left Outer Join CodigoTexto On Codigo = FEmLinea " _
                & ", Ciudad CiuOrigen, Ciudad CiuDestino, AgenciaTransporte, Contenedor" _
                & " Where FEmID IN (" & sIn & ")"
    Else
    
        If Flete > 0 Then
            Cons = "Select FEmID, FEmFAPartir, FEMFHasta, FEmImporte, CiuOrigen.CiuNombre as Origen, CiuDestino.CiuNombre as Destino, ATrNombre, ConAbreviacion, Texto, FEmComentario, FEmDiasLibre, FEmGastosMvd, FEmDiasTransito  " _
                & " From FleteEmbarque " _
                    & " Left Outer Join CodigoTexto On Codigo = FEmLinea " _
                & ", Ciudad CiuOrigen, Ciudad CiuDestino, AgenciaTransporte, Contenedor" _
                & " Where FEmID = " & Flete
        Else
        
            If cAgencia.ListIndex = -1 And cAgencia.Text <> "" Then
                MsgBox "La agencia ingresada no es correcta.", vbExclamation, "Atención"
                cAgencia.SetFocus
                Exit Sub
            End If
            
            If cDestino.ListIndex = -1 And cDestino.Text <> "" Then
                MsgBox "El destino ingresado no es correcto.", vbExclamation, "Atención"
                cDestino.SetFocus
                Exit Sub
            End If

            If cOrigen.ListIndex = -1 And cOrigen.Text <> "" Then
                MsgBox "El origen ingresado no es correcto.", vbExclamation, "Atención"
                cOrigen.SetFocus
                Exit Sub
            End If
            
            If cContenedor.ListIndex = -1 And cContenedor.Text <> "" Then
                MsgBox "El contenedor ingresado no es correcto.", vbExclamation, "Atención"
                cContenedor.SetFocus
                Exit Sub
            End If
        
            If cLinea.ListIndex = -1 And cLinea.Text <> "" Then
                MsgBox "La línea ingresada no es correcta.", vbExclamation, "Atención"
                cLinea.SetFocus
                Exit Sub
            End If
        
            Cons = "Select FEmID, FEmFAPartir, FEMFHasta, FEmImporte, CiuOrigen.CiuNombre as Origen, CiuDestino.CiuNombre as Destino, ATrNombre, ConAbreviacion, Texto, FEmComentario, FEmDiasLibre, FEmGastosMvd, FEmDiasTransito " _
            & " From FleteEmbarque " _
                & " Left Outer Join CodigoTexto On Codigo = FEmLinea " _
            & " , Ciudad CiuOrigen, Ciudad CiuDestino, AgenciaTransporte, Contenedor Where FEmID > 0"
            If cOrigen.ListIndex > -1 Then Cons = Cons & " And FEmOrigen =  " & cOrigen.ItemData(cOrigen.ListIndex)
            If cDestino.ListIndex > -1 Then Cons = Cons & " And FEmDestino = " & cDestino.ItemData(cDestino.ListIndex)
            If cAgencia.ListIndex > -1 Then Cons = Cons & " And FEmAgencia = " & cAgencia.ItemData(cAgencia.ListIndex)
            
            If IsDate(tFAPartir.FechaText) Then
                Cons = Cons & " And FEmFAPartir >= '" & Format(tFAPartir.FechaText, "mm/dd/yyyy") & "'"
                If IsDate(tFHasta.FechaText) Then
                    Cons = Cons & " And FEmFHasta >= '" & Format(tFHasta.FechaText, "mm/dd/yyyy") & "'"
                End If
            Else
                If IsDate(tEmbarco.FechaText) Then
                    Cons = Cons & " And FEmFAPartir >= '" & Format(DateAdd("m", -6, tEmbarco.FechaText), "mm/dd/yyyy") & "'"
                    Cons = Cons & " And FEmFHasta >= '" & Format(tEmbarco.FechaText, "mm/dd/yyyy") & "'"
                Else
                    If IsDate(tFHasta.FechaText) Then
                        Cons = Cons & " And FEmFHasta >= '" & Format(tFHasta.FechaText, "mm/dd/yyyy") & "'"
                    End If
                End If
            End If
            If cContenedor.ListIndex > -1 Then Cons = Cons & " And FEmContenedor = " & cContenedor.ItemData(cContenedor.ListIndex)
            If tImporte.Text <> "" Then Cons = Cons & " And FEmImporte = " & CCur(tImporte.Text)
            If cLinea.ListIndex > -1 Then Cons = Cons & " And FEmLinea = " & cLinea.ItemData(cLinea.ListIndex)
        End If
    End If
    vsDato.Redraw = False
    Cons = Cons & " And FEmOrigen = CiuOrigen.CiuCodigo And FEmDestino = CiuDestino.CiuCodigo " _
        & " And FEmAgencia = ATrCodigo And FEmContenedor = ConCodigo " _
            & " Order by FEmFAPartir Desc"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        If IsDate(tEmbarco.FechaText) Then
            'Busco si ya esta insertada una tupla con los mismos datos
            'y válido las fechas.
            If Not IsNull(RsAux!FEmFAPartir) Then
                Dim lCont As Long
                With vsDato
                    lInserto = 0
                    For lCont = 1 To .Rows - 1
                        If Trim(.Cell(flexcpText, lCont, 0)) = Trim(RsAux!ConAbreviacion) _
                            And Trim(.Cell(flexcpText, lCont, 2)) = Trim(RsAux!ATrNombre) _
                            And Trim(.Cell(flexcpText, lCont, 3)) = Trim(RsAux!origen) _
                            And Trim(.Cell(flexcpText, lCont, 4)) = Trim(RsAux!Destino) _
                            And Trim(.Cell(flexcpText, lCont, 6)) = Trim(RsAux!Texto) Then
                        
                            If CDate(RsAux!FEmFAPartir) < tEmbarco.FechaText And RsAux!FEmFAPartir > CDate(.Cell(flexcpText, .Rows - 1, 1)) Then
                                vsDato.RemoveItem lCont
                                lInserto = vsDato.Rows
                            Else
                                lInserto = -1
                            End If
                            Exit For
                        End If
                    Next lCont
                    If CDate(RsAux!FEmFAPartir) <= tEmbarco.FechaText Then
                        If lInserto = 0 Then
                            lInserto = vsDato.Rows
                        End If
                    Else
                        lInserto = -1
                    End If
                End With
            End If
        Else
            lInserto = vsDato.Rows
        End If
        If lInserto > -1 Then
            With vsDato
                .AddItem "", lInserto
                lValor = RsAux!FEmID: .Cell(flexcpData, .Rows - 1, 0) = lValor
                .Cell(flexcpText, .Rows - 1, 0) = Trim(RsAux!ConAbreviacion)
                If Not IsNull(RsAux!FEmFAPartir) Then .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!FEmFAPartir, "dd/mm/yyyy")
                If Not IsNull(RsAux!FEmFHasta) Then .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!FEmFHasta, "dd/mm/yyyy")
                .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!ATrNombre)
                .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!origen)
                .Cell(flexcpText, .Rows - 1, 5) = Trim(RsAux!Destino)
                .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!FEmImporte, "#,##0.00")
                If Not IsNull(RsAux!Texto) Then .Cell(flexcpText, .Rows - 1, 7) = Trim(RsAux!Texto)
                If Not IsNull(RsAux!FEmComentario) Then .Cell(flexcpText, .Rows - 1, 8) = Trim(RsAux!FEmComentario)
                If Not IsNull(RsAux("FEmDiasLibre")) Then .Cell(flexcpText, .Rows - 1, 9) = Trim(RsAux!FEmDiasLibre)
                If Not IsNull(RsAux("FEmGastosMvd")) Then .Cell(flexcpText, .Rows - 1, 10) = Format(RsAux!FEmGastosMvd, "#,##0.00")
                If Not IsNull(RsAux("FEmDiasTransito")) Then .Cell(flexcpText, .Rows - 1, 11) = RsAux!FEmDiasTransito
                
            End With
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    If vsDato.Rows > 1 Then
        vsDato.Select 1, 0
        Botones True, True, True, False, False, tooMenu, Me
    End If
    vsDato.Redraw = True
    Exit Sub
errAB:
    vsDato.Redraw = True
    clsGeneral.OcurrioError "Ocurrió un error al buscar.", Err.Description
End Sub

Private Function ValidoDatos() As Boolean
    ValidoDatos = False
    
    If cOrigen.ListIndex = -1 And listaOrigen.Lista Is Nothing Then
        MsgBox "El origen es obligatorio.", vbExclamation, "ATENCIÓN": cOrigen.SetFocus: Exit Function
    End If
    If cDestino.ListIndex = -1 Then
        MsgBox "El destino es obligatorio.", vbExclamation, "ATENCIÓN": cDestino.SetFocus: Exit Function
    End If
    If cAgencia.ListIndex = -1 Then
        MsgBox "La agencia es obligatoria.", vbExclamation, "ATENCIÓN": cAgencia.SetFocus: Exit Function
    End If
    If cContenedor.ListIndex = -1 Then
        MsgBox "El contenedor es obligatorio.", vbExclamation, "ATENCIÓN": cOrigen.SetFocus: Exit Function
    End If
    If Not IsNumeric(tImporte.Text) Then
        MsgBox "Debe ingresar un importe.", vbExclamation, "ATENCIÓN": tImporte.SetFocus: Exit Function
    End If
    If Trim(tFAPartir.FechaText) <> "" And tFAPartir.FechaValor = "" Then
        MsgBox "La fecha ingresada no es válida.", vbExclamation, "ATENCIÓN": tFAPartir.SetFocus: Exit Function
    End If
    If Trim(tFHasta.FechaText) <> "" And tFHasta.FechaValor = "" Then
        MsgBox "La fecha ingresada no es válida.", vbExclamation, "ATENCIÓN": tFHasta.SetFocus: Exit Function
    End If
    If Trim(tDiasFree.Text) <> "" And Not IsNumeric(tDiasFree.Text) Then
        MsgBox "Debe ingresar un número.", vbExclamation, "Atención": tDiasFree.SetFocus: Exit Function
    End If
    If Trim(tGastosMvd.Text) <> "" And Not IsNumeric(tGastosMvd.Text) Then
        MsgBox "Debe ingresar un importe.", vbExclamation, "Atención": tGastosMvd.SetFocus: Exit Function
    End If

    Dim sLineas As String
    sLineas = listaLinea.ListaCodigos()

    On Error GoTo errQ
    
    Cons = "SELECT FEmFAPartir Desde, FEmFHasta Hasta, FEmImporte Precio, CiuNombre Origen, Texto Línea " _
        & "FROM FleteEmbarque INNER JOIN CodigoTexto On Codigo = FEmLinea INNER JOIN Ciudad ON FEmOrigen = CiuCodigo " _
        & "WHERE FEmOrigen IN (" & listaOrigen.ListaCodigos & ")" _
        & " And FEmDestino = " & cDestino.ItemData(cDestino.ListIndex) _
        & " And FEmAgencia = " & cAgencia.ItemData(cAgencia.ListIndex) _
        & " And FEmContenedor = " & cContenedor.ItemData(cContenedor.ListIndex)
        
    If sLineas <> "" Then Cons = Cons & " And FEmLinea IN(" & sLineas & ")"
    
    Cons = Cons & " AND ((FEmFAPartir <= '" & Format(tFAPartir.FechaText, "yyyy/mm/dd") & "' AND FEmFHasta >= '" & Format(tFAPartir.FechaText, "yyyy/mm/dd") & "')" _
        & " OR (FEmFAPartir >= '" & Format(tFAPartir.FechaText, "yyyy/mm/dd") & "' AND FEmFAPartir <= '" & Format(tFHasta.FechaText, "yyyy/mm/dd") & "'))" _
        & " AND FEmID <> " & idFlete
    Dim rsV As rdoResultset
    Set rsV = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsV.EOF Then
        rsV.Close
        
        MsgBox "Existen registros que cubren total o parcialmente los datos ingresados.", vbInformation, "Posible datos duplicados"
        
        Dim oHelp As New clsListadeAyuda
        oHelp.ActivarAyuda cBase, Cons, 6000, 0, "Rangos solapados"
        Set oHelp = Nothing
        
        If MsgBox("¿Desea continuar almacenando los datos ingresados?", vbQuestion + vbYesNo, "Posible duplicación") = vbNo Then
            Exit Function
        End If
    Else
        rsV.Close
    End If
    
    If idFlete = 0 Then
        Cons = "Select Count(*) " _
                & " From FleteEmbarque " _
                    & " Left Outer Join CodigoTexto On Codigo = FEmLinea " _
                & ", Ciudad CiuOrigen, Ciudad CiuDestino, AgenciaTransporte, Contenedor" _
                & " Where FEmOrigen IN (" & listaOrigen.ListaCodigos & ")" _
                & " And FEmDestino = " & cDestino.ItemData(cDestino.ListIndex) _
                & " And FEmAgencia = " & cAgencia.ItemData(cAgencia.ListIndex) _
                & " And FEmContenedor = " & cContenedor.ItemData(cContenedor.ListIndex) _
                & " And FEmImporte = " & CCur(tImporte.Text) & " And FEmDiasLibre " & IIf(tDiasFree.Text <> "", " = " & tDiasFree.Text, " Is Null")
    
        If tFAPartir.FechaText <> "" Then Cons = Cons & " And FEmFAPartir >= '" & Format(tFAPartir.FechaText, "yyyy/mm/dd") & "'" _
    
        If Not listaLinea Is Nothing Then
            If listaLinea.Lista.Count > 0 Then
                Cons = Cons & " And FEmLinea IN(" & sLineas & ")"
            End If
        End If
            
        Cons = Cons & " And FEmOrigen = CiuOrigen.CiuCodigo And FEmDestino = CiuDestino.CiuCodigo " _
                & " And FEmAgencia = ATrCodigo And FEmContenedor = ConCodigo " _
                & "-- Order by CiuOrigen.CiuNombre"
    
        Set rsV = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsV.EOF Then
            If rsV(0) > 0 Then
                rsV.Close
                MsgBox "Existen datos ingresados con fecha posterior a la ingresada, verifique.", vbInformation, "Atención"
                Exit Function
            End If
        End If
        rsV.Close
    End If

    ValidoDatos = True
Exit Function
errQ:
    clsGeneral.OcurrioError "Error al validar duplicados.", Err.Description, "Fletes duplicados"
End Function

Private Sub txtConjLinea_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyDelete Then Exit Sub
    ProcesoTextoLista txtConjLinea, listaLinea
End Sub

Private Sub txtConjOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyDelete Then Exit Sub
    ProcesoTextoLista txtConjOrigen, listaOrigen
    
End Sub

Private Sub txtTT_GotFocus()
    With txtTT
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTT_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If Trim(txtTT.Text) <> "" Then
            If Not IsNumeric(txtTT.Text) Then
                MsgBox "TT debe ir de 0 a 254 días", vbExclamation, "Posible error"
                With txtTT
                    .SelStart = 0: .SelLength = Len(.Text)
                End With
                Exit Sub
            End If
        End If
        tFAPartir.SetFocus
    End If
End Sub

Private Sub txtTT_LostFocus()
    If Trim(txtTT.Text) <> "" Then
        If Not IsNumeric(txtTT.Text) Then
            MsgBox "TT debe ir de 0 a 254 días", vbExclamation, "Posible error"
            txtTT.Text = ""
            Exit Sub
        ElseIf Val(txtTT.Text) < 0 Or Val(txtTT.Text) > 254 Then
            MsgBox "TT debe ir de 0 a 254 días", vbExclamation, "Posible error"
            txtTT.Text = ""
            Exit Sub
        End If
    End If
End Sub

Private Sub vsDato_DblClick()
    If Not bEdicion And MnuModificar.Enabled Then
        AccionModificar
    End If
End Sub

Private Sub vsDato_SelChange()
    If vsDato.Row >= 1 Then
        Botones True, True, True, False, False, tooMenu, Me
    Else
        Botones True, False, False, False, False, tooMenu, Me
    End If
End Sub

Private Function MantengoGrilla() As String
Dim sIn As String
Dim iQ As Integer
     sIn = ""
     For iQ = vsDato.FixedRows To vsDato.Rows - 1
        If sIn <> "" Then sIn = sIn & ","
        sIn = sIn & vsDato.Cell(flexcpData, iQ, 0)
     Next iQ
    MantengoGrilla = sIn
End Function

Private Sub vsDato_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If Row < vsDato.FixedRows Then Cancel = True: Exit Sub
    Select Case Col
        Case 1, 2
            If IsDate(vsDato.EditText) Then vsDato.EditText = Format(vsDato.EditText, "dd/mm/yyyy")
            If Not IsDate(vsDato.EditText) And vsDato.EditText <> "" Then
                MsgBox "Formato incorrecto, [Esc] = cancela.", vbCritical, "Atención"
                Cancel = True
            End If
            
        Case 6, 9, 10
            If Not IsNumeric(vsDato.EditText) And vsDato.EditText <> "" Then
                MsgBox "Formato incorrecto, [Esc] = cancela.", vbCritical, "Atención"
                Cancel = True
            End If
    End Select
    loc_SaveByGrid Col, vsDato.Cell(flexcpData, Row, 0)
    
End Sub

