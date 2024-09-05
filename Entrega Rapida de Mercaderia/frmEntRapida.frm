VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmEntRapida 
   BackColor       =   &H80000005&
   Caption         =   "Entrega rápida de Mercadería"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6780
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEntRapida.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picGrabo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1425
      ScaleWidth      =   6465
      TabIndex        =   7
      Top             =   840
      Width           =   6495
      Begin VSFlex6DAOCtl.vsFlexGrid vsGrabo 
         Height          =   2775
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   4895
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
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
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   10
         FixedRows       =   0
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
         MergeCells      =   6
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
   Begin VB.Timer tmDisp 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   4320
      Top             =   2640
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsPendiente 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2990
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   285
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
   Begin VB.TextBox tFactura 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6000
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntRapida.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntRapida.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntRapida.frx":09EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sesion"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "refrescar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "sucursal"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "lista"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsPasados 
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3625
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   8421504
      ForeColor       =   16777215
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   8421504
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   270
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
   Begin VB.Label lbTitPasados 
      BackStyle       =   0  'Transparent
      Caption         =   "  Ú&ltimos documentos pasados"
      ForeColor       =   &H00000080&
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   3495
   End
   Begin VB.Label lbUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Entrega:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Documento:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   975
   End
End
Attribute VB_Name = "frmEntRapida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tArtSonido
    ID As Long
    Sonido As String
End Type
Private arrSonido() As tArtSonido

Private Enum TipoLocal
    Camion = 1
    Deposito = 2
End Enum

Private Type tLista
    ID As Long
    Nombre As String
    Articulos As String
End Type

Private sWav As String

Private sDocHide As String          'Documentos ocultos x el usuario.
Private tListArt As tLista              'Lista de artículos a seleccionar en los documentos.
Private iSuc As Integer                'Sucursal de la cual busco los documentos.

Private Type tUID
    Codigo As Long
    Identificacion As String
End Type
Private arrUID(12) As tUID
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Function fnc_FindLista() As Boolean
On Error GoTo errFL
Dim objH As New clsListadeAyuda
    tListArt.ID = 0: tListArt.Nombre = "": tListArt.Articulos = ""
    With objH
        If .ActivarAyuda(cBase, "Select LERCodigo, LERArticulos, LERNombre as Nombre From ListaArtsEntRapida Order By LERNombre", 4000, 2, "Listas") > 0 Then
            tListArt.ID = .RetornoDatoSeleccionado(0)
            tListArt.Nombre = .RetornoDatoSeleccionado(2)
            tListArt.Articulos = .RetornoDatoSeleccionado(1)
        End If
    End With
    Set objH = Nothing
    fnc_FindLista = (tListArt.ID <> 0)
    Me.Caption = "Entrega rápida de Mercadería (" & tListArt.Nombre & ")"
    If tListArt.ID = 0 Then
        Me.BackColor = &H8080DD
    Else
        Me.BackColor = vbWindowBackground
    End If
Exit Function
errFL:
    objG.OcurrioError "Error al buscar la lista.", Err.Description, "Lista"
End Function

Private Sub loc_CargoListaSeleccionada(ByVal idLista As Long)
On Error GoTo errCLS
Dim rsL As rdoResultset
    With tListArt
        .ID = 0: .Articulos = "": .Nombre = ""
    End With
    Set rsL = cBase.OpenResultset("Select * From ListaArtsEntRapida Where LERCodigo =" & idLista, rdOpenDynamic, rdConcurValues)
    If Not rsL.EOF Then
        With tListArt
            .ID = rsL("LERCodigo")
            .Nombre = Trim(rsL("LERNombre"))
            .Articulos = Trim(rsL("LERArticulos"))
        End With
    End If
    rsL.Close
    Me.Caption = "Entrega rápida de Mercadería (" & tListArt.Nombre & ")"
Exit Sub
errCLS:
    objG.OcurrioError "Error al cargar la lista.", Err.Description, "Error (cargolistaseleccionada)"
End Sub

Private Sub loc_ArtSonidos()
On Error GoTo errAS
Dim rsAF As rdoResultset
    Erase arrSonido
    ReDim arrSonido(0)
    Set rsAF = cBase.OpenResultset("Select AFaArticulo, AFaSonido From ArticuloFacturacion Where AFaArticulo In(" & tListArt.Articulos & ") And AFaSonido Is Not Null", rdOpenDynamic, rdConcurValues)
    Do While Not rsAF.EOF
        ReDim Preserve arrSonido(UBound(arrSonido) + 1)
        With arrSonido(UBound(arrSonido))
            .ID = rsAF("AFaArticulo")
            .Sonido = Trim(rsAF("AFaSonido"))
        End With
        rsAF.MoveNext
    Loop
    rsAF.Close
Exit Sub
errAS:
    objG.OcurrioError "Error al cargar los sonidos de los artículos de la lista.", Err.Description
End Sub

Private Sub loc_SetSonidoArticulo(ByVal iArt As Long)
Dim iQ As Integer
    For iQ = 1 To UBound(arrSonido)
        If iArt = arrSonido(iQ).ID Then
            If arrSonido(iQ).Sonido <> "" Then
                loc_SetSonido arrSonido(iQ).Sonido
                me_Wait
            End If
            Exit Sub
        End If
    Next
End Sub

Private Sub loc_SetGrabo(ByVal bShow As Boolean, Optional iStart As Integer, Optional iDocumento As Long)
Dim iQ As Integer
On Error Resume Next
    If bShow Then
        With vsGrabo
            .Rows = 0
            .ColWidth(0) = .ClientWidth - 1000
            .HighLight = flexHighlightNever
            .AddItem vsPendiente.Cell(flexcpText, iStart, 2)
            .Cell(flexcpFontSize, .Rows - 1, 0) = 16
            .RowHeight(.Rows - 1) = 600
            .Cell(flexcpBackColor, .Rows - 1, 0, , 1) = &HD0F0FF
            .CellAlignment = flexAlignCenterCenter
            
            
            For iQ = iStart To vsPendiente.Rows - 1
                If iDocumento = vsPendiente.Cell(flexcpData, iQ, 0) Then
                    .AddItem vsPendiente.Cell(flexcpText, iQ, 4)
                    .Cell(flexcpFontSize, .Rows - 1, 0, , 1) = 26
                    .RowHeight(.Rows - 1) = 700
                    .CellAlignment = flexAlignCenterTop
                    
                    If vsPendiente.Cell(flexcpValue, iStart, 3) > 1 Then
                        .Cell(flexcpText, .Rows - 1, 1) = vsPendiente.Cell(flexcpValue, iQ, 3)
                        .Cell(flexcpBackColor, .Rows - 1, 1) = &HC0&
                        .Cell(flexcpForeColor, .Rows - 1, 1) = vbWhite
                    End If
                    .Cell(flexcpAlignment, .Rows - 1, 1) = flexAlignCenterCenter
                End If
            Next
        End With
        picGrabo.Visible = True
        tFactura.Enabled = False
    Else
        picGrabo.Visible = False
        vsGrabo.Rows = 0
        tFactura.Enabled = True
    End If
    Me.Refresh
End Sub

Private Sub loc_CambioColores()
On Error Resume Next
Dim iQ As Integer, iAnt As Long, iColor As Long
    iAnt = 0
    With vsPendiente
        For iQ = .FixedRows To .Rows - 1
            If Not .RowHidden(iQ) Then
                If iAnt <> .Cell(flexcpData, iQ, 0) Then
                    iColor = IIf(iColor = vbWindowBackground, "&HCDFAFA", vbWindowBackground)
                    iAnt = .Cell(flexcpData, iQ, 0)
                End If
                .Cell(flexcpBackColor, iQ, 0, iQ, .Cols - 1) = iColor
            End If
        Next
        .Refresh
    End With
End Sub

Private Sub me_Wait()
Dim iQ As Integer
    For iQ = 1 To 2000
        DoEvents
    Next
End Sub

Private Sub loc_SetSonido(ByVal sFile As String)
On Error Resume Next
Dim Result As Long
    Result = sndPlaySound(sWav & sFile, 1)
End Sub

Private Sub loc_APasados(ByVal iDoc As Long, ByVal iStart As Integer)
Dim iQ As Integer, iQ1 As Long
        
    On Error Resume Next
    
    With vsPasados
        If .Rows > 12 Then
            iQ1 = .Cell(flexcpData, .Rows - 1, 0)
            For iQ = .Rows - 1 To .FixedRows Step -1
                If iQ1 = CLng(.Cell(flexcpData, iQ, 0)) Then
                    .RemoveItem iQ
                End If
            Next
        End If
        
        For iQ = .FixedRows To .Rows - 1
            .RowHeight(iQ) = 270
        Next
    End With
    
    With vsPendiente
        iQ1 = 0
        For iQ = iStart To .Rows - 1
            If iDoc = .Cell(flexcpData, iQ, 0) Then
                iQ1 = iQ1 + 1
                vsPasados.AddItem .Cell(flexcpText, iQ, 0), 1
                vsPasados.Cell(flexcpData, 1, 0) = .Cell(flexcpData, iQ, 0)
                vsPasados.Cell(flexcpText, 1, 1) = .Cell(flexcpText, iQ, 3)
                vsPasados.Cell(flexcpText, 1, 2) = .Cell(flexcpText, iQ, 4)
                vsPasados.Cell(flexcpText, 1, 3) = .Cell(flexcpText, iQ, 2)
            End If
        Next
    End With
    
    With vsPasados
        .Cell(flexcpBackColor, .FixedRows, 0, .Rows - 1, .Cols - 1) = &H808080
        .Cell(flexcpBackColor, .FixedRows, 0, iQ1, .Cols - 1) = &H8080&
        .Cell(flexcpFontBold, .FixedRows, 0, .Rows - 1, .Cols - 1) = False
        .Cell(flexcpFontBold, .FixedRows, 0, iQ1, .Cols - 1) = True
    End With
    
    
End Sub

Private Function fnc_SelectCodigoInGrid(ByVal iCod As Long) As Integer
On Error GoTo errSCI
Dim iQ As Integer
Dim iStart As Integer, iEnd As Integer

    With vsPendiente
        For iQ = .FixedRows To .Rows - 1
            If iCod = .Cell(flexcpData, iQ, 0) Then
                If iStart = 0 Then iStart = iQ
                iEnd = iQ
            Else
                If iStart > 0 Then Exit For
            End If
        Next
        If iStart > 0 Then
            .Select 1, 0, 1, 0
            .Select iStart, 0, iEnd, .Cols - 1
            .BackColorSel = &H80&
            iQ = .CellTop
        End If
    End With
    fnc_SelectCodigoInGrid = iStart
errSCI:
End Function

Private Sub loc_BuscoDocumento(ByVal iDocCodigo As Long)
Dim sPor As String, sQuery As String
Dim rsG As rdoResultset
Dim iIndex As Integer, iAux As Long
Dim iTipo As Integer
    On Error GoTo errFD
    
    If tListArt.ID = 0 Then Exit Sub
    
    sQuery = "Select DocCodigo From Documento, Renglon " & _
        " Where DocCodigo = " & iDocCodigo & _
        " And DocFecha BetWeen '" & Format(Date, "yyyy/mm/dd 00:00:00") & "' And '" & Format(Date, "yyyy/mm/dd 23:59:59") & "'" & _
        " And RenARetirar > 0 And RenArticulo Not In(" & tListArt.Articulos & ") And DocCodigo = RenDocumento"
    
    sQuery = "Select DocCodigo, DocTipo, rtrim(DocSerie) as DS, DocNumero, DocFecha, DocFModificacion, CliCodigo, CliCiRuc, CliTipo, " & _
                "NPer = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2), " & _
                "NEmp = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')'), ArtID, rTrim(ArtNombre) as AN, ArtNroSerie, RenARetirar " & _
                " From Documento, Renglon, Articulo, Cliente " & _
                    " Left Outer Join CPersona On CliCodigo = CPeCliente " & _
                    " Left Outer Join CEmpresa On CliCodigo = CEmCliente " & _
                " Where DocCodigo = " & iDocCodigo & _
                " And DocFecha BetWeen '" & Format(Date, "yyyy/mm/dd 00:00:00") & "' And '" & Format(Date, "yyyy/mm/dd 23:59:59") & "'" & _
                " And RenARetirar > 0 And RenArticulo In(" & tListArt.Articulos & ")" & _
                " And DocCodigo Not In (" & sQuery & ") And DocAnulado = 0 " & _
                " And DocSucursal = 5 And DocPendiente Is Null And DocTipo IN (1, 2)" & _
                " And DocCodigo = RenDocumento And RenArticulo = ArtID And DocCliente = CliCodigo Order by DocCodigo"
                
    Set rsG = cBase.OpenResultset(sQuery, rdOpenDynamic, rdConcurValues)

    If Not rsG.EOF Then
        iDocCodigo = rsG("DocCodigo")
        iIndex = fnc_SelectCodigoInGrid(iDocCodigo)
        iTipo = rsG("DocTipo")
        
        If iIndex = 0 Then
        'No existe --> lo inserto.
            iIndex = vsPendiente.Rows
            Do While Not rsG.EOF
                With vsPendiente
                    .AddItem rsG("DS") & " " & rsG("DocNumero")
                    .Cell(flexcpText, .Rows - 1, 1) = Format(rsG("DocFecha"), "hh:nn") & " (" & Abs(DateDiff("n", Now, rsG("DocFecha"))) & ")"
                    If rsG("CliTipo") = 1 Then sQuery = rsG("NPer") Else sQuery = rsG("NEmp")
                    .Cell(flexcpText, .Rows - 1, 2) = sQuery
                    .Cell(flexcpText, .Rows - 1, 3) = rsG("RenARetirar")
                    If rsG("RenARetirar") > 1 Then .Cell(flexcpForeColor, .Rows - 1, 3) = &H80&: .Cell(flexcpFontBold, .Rows - 1, 3) = True
                    .Cell(flexcpText, .Rows - 1, 4) = rsG("AN")
                    
                    iAux = rsG("ArtID"): .Cell(flexcpData, .Rows - 1, 3) = iAux
                    iAux = rsG("DocCodigo"): .Cell(flexcpData, .Rows - 1, 0) = iAux
                    sQuery = rsG("DocFModificacion"): .Cell(flexcpData, .Rows - 1, 1) = sQuery
                    If rsG("ArtNroSerie") Then .Cell(flexcpData, .Rows - 1, 2) = 1
                End With
                rsG.MoveNext
            Loop
            With vsPendiente
                .BackColorSel = &H80&: .Select iIndex, 0, .Rows - 1, .Cols - 1
            End With
        Else
            If CDate(vsPendiente.Cell(flexcpData, iIndex, 1)) <> rsG("DocFModificacion") Then
                loc_SetSonido "entregamal.wav"
                MsgBox "El documento fue modificado no podrá darlo entregado.", vbExclamation, "Posible error"
                iIndex = -1
            End If
        End If
        
        If iIndex > 0 Then
            'PASO A DAR POR ENTREGADO

'            If MsgBox("¿Confirma dar por entregada la mercadería?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
                tFactura.Enabled = False
                loc_SetGrabo True, iIndex, iDocCodigo
                If fnc_SaveEntrega(iDocCodigo, CDate(vsPendiente.Cell(flexcpData, iIndex, 1)), iTipo, iIndex) Then
                    loc_APasados iDocCodigo, iIndex
                    loc_SetSonido "entregaok.wav"
                Else
                    loc_SetSonido "entregamal.wav"
                End If
                tFactura.Text = ""
                tmDisp.Enabled = True
                tmDisp.Tag = "2"
'            End If
        End If
    Else
        loc_SetSonido "entregamal.wav"
        MsgBox "El documento ingresado no existe o no es válido para entrega rápida.", vbExclamation, "Atención"
    End If
    rsG.Close

errResume:
    loc_FillGrid
    Exit Sub
    
errFD:
    tFactura.Text = ""
    loc_SetSonido "entregamal.wav"
    objG.OcurrioError "Error al validar el documento.", Err.Description, "Valido Entrega"
    loc_SetGrabo False
    Resume errResume
End Sub

Private Sub loc_GetNroSerie(ByVal iCodDoc As Long, ByVal iStart As Integer)
Dim iQ As Integer
Dim aNroSerie As String
    With vsPendiente
        For iQ = iStart To .Rows - 1
            If Val(.Cell(flexcpData, iQ, 0)) = iCodDoc Then
                If Val(.Cell(flexcpData, iQ, 2)) = 1 And Val(.Cell(flexcpText, iQ, 3)) = 1 Then
                    Do While aNroSerie = ""
                        aNroSerie = InputBox("Ingrese el número de nerie del artículo entregado.", .Cell(flexcpText, iQ, 4))
                        If Trim(aNroSerie) <> "" Then
                            .Cell(flexcpText, iQ, .Cols - 1) = aNroSerie
                        End If
                    Loop
                End If
            End If
        Next
    End With
End Sub

Private Function fnc_SaveEntrega(ByVal iDoc As Long, ByVal gFechaDocumento As Date, ByVal gTipoDoc As Byte, ByVal iStart As Integer) As Boolean
Dim aTexto As String, Cons As String, rsAux As rdoResultset
   
    loc_GetNroSerie iDoc, iStart
    FechaDelServidor
    gFechaServidor = Now
    On Error GoTo errorBT
    cBase.BeginTrans
    On Error GoTo errRoll
    
    Cons = "Select * from Documento Where DocCodigo = " & iDoc
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If rsAux!DocFModificacion <> gFechaDocumento Then
        aTexto = "El documento seleccionado ha sido modificado por otra terminal. Vuelva a consultar."
        rsAux.Close
        GoTo errorET
        Exit Function
    Else
        rsAux.Edit
        rsAux!DocFModificacion = Format(Now, "yyyy/mm/dd hh:nn:ss")
        rsAux.Update
    End If
    rsAux.Close '-----------------------------------------------------------------------------------------------------------
    
    'Grabo Artículos.......................................................................................................
    loc_GraboRenglon iDoc, gTipoDoc, iStart
    loc_GraboProductosVendidos iStart, iDoc
    '...........................................................................................................................
    
    cBase.CommitTrans
    fnc_SaveEntrega = True
Exit Function
errorBT:
    Screen.MousePointer = 0
    objG.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación."
    Exit Function
errorET:
    Resume errRoll
errRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    If aTexto = "" Then aTexto = "No se ha podido realizar la transacción. Reintente la operación."
    objG.OcurrioError aTexto
    Exit Function

End Function

Private Sub loc_GraboRenglon(ByVal iCodDoc As Long, ByVal iTipoDoc As Integer, ByVal iStart As Integer)
Dim iQ As Integer
Dim rsR As rdoResultset
Dim Cons As String

    With vsPendiente
        For iQ = iStart To .Rows - 1
            If Val(.Cell(flexcpData, iQ, 0)) = iCodDoc Then
            
                'Ejecuto sonido para esta artículo.
                loc_SetSonidoArticulo .Cell(flexcpData, iQ, 3)
                
                Cons = "Select * From Renglon Where RenDocumento = " & iCodDoc & _
                            " And RenArticulo = " & .Cell(flexcpData, iQ, 3)
                Set rsR = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                rsR.Edit
                rsR("RenARetirar") = 0
                rsR.Update
                rsR.Close
                                    
                'Marco la Baja del STOCK AL LOCAL
                'Genero Movimiento
                MarcoMovimientoStockFisico arrUID(lbUsuario.Tag).Codigo, TipoLocal.Deposito, paCodigoDeSucursal, .Cell(flexcpData, iQ, 3), .Cell(flexcpValue, iQ, 3), paEstadoArticuloEntrega, -1, iTipoDoc, iCodDoc
                'Bajo del Stock en Local
                MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, .Cell(flexcpData, iQ, 3), .Cell(flexcpValue, iQ, 3), paEstadoArticuloEntrega, -1
                
                'Marco el Movimiento del STOCK VIRTUAL
                'Genero Movimiento
                MarcoMovimientoStockEstado arrUID(lbUsuario.Tag).Codigo, .Cell(flexcpData, iQ, 3), .Cell(flexcpValue, iQ, 3), TipoMovimientoEstado.ARetirar, -1, iTipoDoc, iCodDoc, paCodigoDeSucursal
                'Bajo del Stock Total
                MarcoMovimientoStockTotal .Cell(flexcpData, iQ, 3), TipoEstadoMercaderia.Virtual, TipoMovimientoEstado.ARetirar, CCur(.Cell(flexcpValue, iQ, 3)), -1
            End If
        Next
    End With

End Sub

Private Sub loc_GraboProductosVendidos(ByVal iStart As Integer, ByVal idDocumento As Long)
Dim iQ As Integer
Dim Cons As String
Dim rsPV As rdoResultset

    Cons = "Select * from ProductosVendidos Where PVeDocumento = " & idDocumento
    Set rsPV = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    With vsPendiente
        For iQ = iStart To .Rows - 1
            If idDocumento = .Cell(flexcpData, iQ, 0) And .Cell(flexcpText, iQ, 5) <> "" Then
                rsPV.AddNew
                rsPV!PVeDocumento = idDocumento
                rsPV!PVeArticulo = .Cell(flexcpData, iQ, 0)
                rsPV!PVeNSerie = .Cell(flexcpText, iQ, 5)
                rsPV.Update
            End If
        Next
    End With
    rsPV.Close
End Sub

Private Sub FechaDelServidor()
On Error GoTo errFs
    Dim RsF As rdoResultset
    Set RsF = cBase.OpenResultset("Select GetDate()", rdOpenDynamic, rdConcurValues)
    Date = RsF(0): Time = RsF(0)
    RsF.Close
errFs:
End Sub

'----------------------------------------------------------------------------------
'   Interpreta el Texto del Codigo de Barras
'   Formato:    XDXXXX          TipoDocumento   D Numero de Documento
'----------------------------------------------------------------------------------
Private Sub loc_FormatoBarras(Texto As String)
Dim aCodDoc As Long, gTipo As Byte
    
    On Error GoTo errInt
    
    If tListArt.ID = 0 Then
        MsgBox "Debe seleccionar una lista.", vbExclamation, "Atención"
        Me.BackColor = &H8080DD
        Exit Sub
    Else
        Me.BackColor = vbWindowBackground
    End If
    
    
    Texto = UCase(Texto)
    
    '1) Veo si es x codigo de barras o x ids de documento
    If (Mid(Texto, 2, 1) = "D" And IsNumeric(Mid(Texto, 1, 1)) And Len(Texto) > 3) Or (Mid(Texto, 3, 1) = "D" And IsNumeric(Mid(Texto, 1, 2)) And Len(Texto) > 4) Then   'Codigo de Barras
        gTipo = CLng(Mid(Texto, 1, InStr(Texto, "D") - 1))
        aCodDoc = CLng(Trim(Mid(Texto, InStr(Texto, "D") + 1, Len(Texto))))
    Else
        'Puso Serie y Numero de Documento o Numero de Remito
        If IsNumeric(Texto) Then
            gTipo = TipoDocumento.Remito        'Remito
            aCodDoc = Texto
        Else
            'Puso Serie y Numero de Documento
             If Not modPersistencia.Documento_BuscoDocPorTexto(Texto, aCodDoc, gTipo) Then aCodDoc = -1
        End If
        
    End If
    
    If aCodDoc = -1 Then
        MsgBox "No existe un documento que coincida con los valores ingresados.", vbExclamation, "No hay Datos"
        Exit Sub
    End If
    
    Select Case gTipo
'        Case TipoDocumento.Remito:  BuscoRemito aCodDoc
        Case TipoDocumento.Contado, TipoDocumento.Credito: loc_BuscoDocumento aCodDoc
'        Case TipoDocumento.NotaCredito, TipoDocumento.NotaDevolucion, TipoDocumento.NotaEspecial: BuscoNota Tipo:=gTipo, Codigo:=aCodDoc
        Case Else
            loc_SetSonido False
            MsgBox "El código de barras ingresado no es correcto. El documento no coincide con los predefinidos.", vbCritical, "ATENCIÓN"
    End Select
    Screen.MousePointer = 0
    Exit Sub
    
errInt:
    Screen.MousePointer = 0
    objG.OcurrioError "Error al interpretar el código de barras.", Err.Description
End Sub

Private Sub loc_CargoParametros()
On Error GoTo errCL
Dim rsL As rdoResultset
    Set rsL = cBase.OpenResultset("Select * From Parametro " & _
                                        "Where ParNombre IN('Provisorio', 'EstadoArticuloEntrega')", _
                                        rdOpenDynamic, rdConcurValues)
    Do While Not rsL.EOF
        Select Case LCase(Trim(rsL("ParNombre")))
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = rsL!ParValor
'            Case "provisorio": sListArt = Trim(rsL("ParTexto"))
        End Select
        rsL.MoveNext
    Loop
    rsL.Close
Exit Sub
errCL:
    objG.OcurrioError "Error al obtener la lista.", Err.Description, "Cargo Lista"
End Sub

Private Sub loc_HideDocumento(ByVal iDoc As Long)
On Error Resume Next
Dim iQ As Integer
    sDocHide = sDocHide & IIf(sDocHide <> "", ", ", "") & iDoc
    With vsPendiente
        For iQ = .FixedRows To .Rows - 1
            If .Cell(flexcpData, iQ, 0) = iDoc Then .RowHidden(iQ) = True
        Next
    End With
End Sub

Private Sub loc_SetUsuario(ByVal iIndex As Byte)
    With arrUID(iIndex)
        lbUsuario.Caption = "Entrega: " & .Identificacion
        lbUsuario.Tag = iIndex
    End With
End Sub

Private Sub loc_GetSesion(Optional iIndex As Byte = 0)
    With InUsuario
        .pUsuarioCodigo = 0
        .pUsuarioNombre = ""
        .pUsuarioTecla = iIndex
        .Show vbModal
        If .pUsuarioCodigo > 0 Then
            If arrUID(.pUsuarioTecla).Codigo > 0 Then MsgBox "La tecla estaba utilizada, avise el cambio.", vbInformation, "Atención"
            
            arrUID(.pUsuarioTecla).Codigo = .pUsuarioCodigo
            arrUID(.pUsuarioTecla).Identificacion = .pUsuarioNombre
            
            loc_SetUsuario .pUsuarioTecla
        End If
    End With
End Sub

Private Sub loc_FillGrid()
On Error GoTo errFG
Dim sQuery As String
Dim rsG As rdoResultset
Dim iAux As Long, iLast As Long, iColor As Long

    If tListArt.ID = 0 Then
        Me.BackColor = &H8080DD
    Else
        Me.BackColor = vbWindowBackground
    End If

    With vsPendiente
        .Rows = .FixedRows
        .BackColorSel = &H8000000D
        If tListArt.ID = 0 Then Exit Sub
        .Redraw = False
    End With
    
    
    
    iColor = 0
    
    sQuery = "Select DocCodigo From Documento, Renglon " & _
        " Where DocFecha BetWeen '" & Format(Date, "yyyy/mm/dd 00:00:00") & "' And '" & Format(Date, "yyyy/mm/dd 23:59:59") & "'" & _
        " And RenARetirar > 0 And RenArticulo Not In(" & tListArt.Articulos & ") And DocCodigo = RenDocumento"
    
    sQuery = "Select DocCodigo, rtrim(DocSerie) as DS, DocNumero, DocFecha, DocFModificacion, CliCodigo, CliCiRuc, CliTipo, " & _
                "NPer = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2), " & _
                "NEmp = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')'), ArtID, ArtCodigo, rTrim(ArtNombre) as AN, ArtNroSerie, RenARetirar " & _
                " From Documento, Renglon, Articulo, Cliente " & _
                    " Left Outer Join CPersona On CliCodigo = CPeCliente " & _
                    " Left Outer Join CEmpresa On CliCodigo = CEmCliente " & _
                " Where DocFecha BetWeen '" & Format(Date, "yyyy/mm/dd 00:00:00") & "' And '" & Format(Date, "yyyy/mm/dd 23:59:59") & "'" & _
                " And RenARetirar > 0 And RenArticulo In(" & tListArt.Articulos & ")" & _
                " And DocTipo IN (1, 2) And DocCodigo Not In (" & sQuery & ") And DocAnulado = 0 " & _
                " And DocCodigo Not In (" & IIf(sDocHide = "", "0", sDocHide) & ") And DocSucursal = 5 And DocPendiente Is Null" & _
                " And DocCodigo = RenDocumento And RenArticulo = ArtID And DocCliente = CliCodigo Order by DocCodigo"
                
    Set rsG = cBase.OpenResultset(sQuery, rdOpenDynamic, rdConcurValues)
    
    Do While Not rsG.EOF
        With vsPendiente
            .AddItem rsG("DS") & " " & rsG("DocNumero")
            .Cell(flexcpText, .Rows - 1, 1) = Format(rsG("DocFecha"), "hh:nn") & " (" & Abs(DateDiff("n", Now, rsG("DocFecha"))) & ")"
            If rsG("CliTipo") = 1 Then sQuery = rsG("NPer") Else sQuery = rsG("NEmp")
            .Cell(flexcpText, .Rows - 1, 2) = sQuery
            .Cell(flexcpText, .Rows - 1, 3) = rsG("RenARetirar")
            If rsG("RenARetirar") > 1 Then .Cell(flexcpForeColor, .Rows - 1, 3) = &H80&: .Cell(flexcpFontBold, .Rows - 1, 3) = True
            .Cell(flexcpText, .Rows - 1, 4) = rsG("AN")
            
            iAux = rsG("ArtID"): .Cell(flexcpData, .Rows - 1, 3) = iAux
            iAux = rsG("DocCodigo"): .Cell(flexcpData, .Rows - 1, 0) = iAux
            If rsG("ArtNroSerie") Then .Cell(flexcpData, .Rows - 1, 2) = 1
            sQuery = rsG("DocFModificacion"): .Cell(flexcpData, .Rows - 1, 1) = sQuery
            
            If iAux <> iLast Then
                iColor = IIf(iColor = vbWindowBackground, "&HCDFAFA", vbWindowBackground)
                iLast = iAux
            End If
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = iColor
        End With
        rsG.MoveNext
    Loop
    rsG.Close
    vsPendiente.Redraw = True
    On Error Resume Next
    If tFactura.Enabled Then tFactura.SetFocus
    
Exit Sub
errFG:
    vsPendiente.Redraw = True
    objG.OcurrioError "Error al cargar la grilla.", Err.Description, "Cargar Grilla"
End Sub

Private Sub loc_InitForm()
    picGrabo.Visible = False
    With vsGrabo
        .ExtendLastCol = True
        .Cols = 2
        .ColWidth(0) = 3000
        .ColWidth(1) = 1000
    End With
    With vsPendiente
        .Rows = .FixedRows
        .Cols = 0
        .FormatString = "Factura|Mínutos|Cliente|>Q|Artículo|#Serie"
        .ExtendLastCol = True
        .ColWidth(0) = 1000: .ColWidth(1) = 1100: .ColWidth(2) = 3000: .ColWidth(3) = 400: .ColWidth(4) = 100
        .ColHidden(.Cols - 1) = True
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(0) = True ': .MergeCol(2) = True
    End With
    With vsPasados
        .Rows = .FixedRows
        .Cols = 0
        .FormatString = "Factura|>Q|Artículo|Cliente"
        .ExtendLastCol = True
        .ColWidth(0) = 1000: .ColWidth(1) = 400: .ColWidth(2) = 3500: .ColWidth(3) = 400
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(0) = True ': .MergeCol(2) = True
    End With
    
End Sub

Private Sub Form_Load()
On Error GoTo errL
    
    ObtengoSeteoForm Me
    loc_InitForm
    loc_CargoParametros
    Dim sID As String
    
    sID = GetSetting(App.Title, "Settings", "ListaArtEnt", "")
    If sID <> "" Then
        loc_CargoListaSeleccionada (Val(sID))
    End If
    
    If tListArt.ID = 0 Then fnc_FindLista
        
    If tListArt.ID = 0 Then
        MsgBox "Atención no existe una lista de artículos seleccionada.", vbExclamation, "Atención"
    Else
        loc_ArtSonidos
    End If
    'pido el primer usuario
    loc_GetSesion
    
    On Error Resume Next
    ChDir App.Path
    ChDir ("..")
    If Dir(CurDir & "\Sonidos", vbDirectory) <> "" Then sWav = CurDir & "\Sonidos\"
    
    loc_FillGrid
Exit Sub
errL:
    objG.OcurrioError "Error al cargar el formulario.", Err.Description, "Load"
End Sub

Private Sub Form_Resize()
On Error Resume Next
    vsPasados.Move 0, Me.ScaleHeight - vsPasados.Height, Me.ScaleWidth
    lbTitPasados.Move 0, vsPasados.Top - lbTitPasados.Height
    vsPendiente.Move 0, vsPendiente.Top, ScaleWidth, lbTitPasados.Top - vsPendiente.Top - 15
    picGrabo.Move 0, vsPendiente.Top, vsPendiente.Width, vsPendiente.Height
    vsGrabo.Move 0, 0, picGrabo.ScaleWidth, picGrabo.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    'Guardo la lista que esté en uso.
    SaveSetting App.Title, "Settings", "ListaArtEnt", tListArt.ID
    Set objG = Nothing
    Set oUsers = Nothing
    CierroConexion
    GuardoSeteoForm Me
End Sub

Private Sub tFactura_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift <> 0 Then Exit Sub
On Error GoTo errKD
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF9, vbKeyF11, vbKeyF12
            If arrUID(KeyCode - (vbKeyF1 - 1)).Codigo > 0 Then
                loc_SetUsuario KeyCode - (vbKeyF1 - 1)
            Else
                loc_GetSesion KeyCode - (vbKeyF1 - 1)
            End If
    End Select
Exit Sub
errKD:
    objG.OcurrioError "Error inesperado.", Err.Description, "Error (textbox keydown)"
End Sub

Private Sub tFactura_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Val(lbUsuario.Tag) = 0 Then
            tFactura.Text = ""
            loc_GetSesion
            Exit Sub
        End If
        If Trim(tFactura.Text) <> "" Then loc_FormatoBarras tFactura.Text
    End If

End Sub

Private Sub tmDisp_Timer()
    tmDisp.Enabled = False
    If Val(tmDisp.Tag) = 2 Then
        'oculto picture con grid.
        loc_SetGrabo False
    End If
    On Error Resume Next
    tFactura.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case LCase(Button.Key)
        Case "sesion": loc_GetSesion
        Case "refrescar": loc_FillGrid
        Case "lista": fnc_FindLista
    End Select
End Sub

Private Sub vsPendiente_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errKP
    
    If Shift <> 0 Then Exit Sub
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF9, vbKeyF11, vbKeyF12
            If Val(lbUsuario.Tag) <> KeyCode - (vbKeyF1 - 1) Then
                If arrUID(KeyCode - (vbKeyF1 - 1)).Codigo > 0 Then
                    loc_SetUsuario KeyCode - (vbKeyF1 - 1)
                Else
                    loc_GetSesion KeyCode - (vbKeyF1 - 1)
                End If
            End If
            tFactura.SetFocus
        Case vbKeyDelete
            tmDisp.Enabled = False
            If vsPendiente.Row >= vsPendiente.FixedRows Then
                loc_HideDocumento vsPendiente.Cell(flexcpData, vsPendiente.Row, 0)
                loc_CambioColores
            End If
            tmDisp.Enabled = True
            tmDisp.Tag = 1
        
        Case vbKeyReturn
            On Error Resume Next
            tFactura.SetFocus
    End Select
Exit Sub
errKP:
    objG.OcurrioError "Error inesperado.", Err.Description, "Error (Grilla keydown)"
End Sub

