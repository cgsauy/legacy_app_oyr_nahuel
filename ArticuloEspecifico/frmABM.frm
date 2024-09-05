VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{923DD7D8-A030-4239-BCD4-51FDB459E0FE}#4.0#0"; "orgComboCalculator.ocx"
Object = "{B443E3A5-0B4D-4B43-B11D-47B68DC130D7}#1.7#0"; "orArticulo.ocx"
Object = "{190700F0-8894-461B-B9F5-5E731283F4E1}#1.1#0"; "orHiperlink.ocx"
Begin VB.Form frmABM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artículos específicos"
   ClientHeight    =   5685
   ClientLeft      =   2895
   ClientTop       =   2235
   ClientWidth     =   7245
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmABM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7245
   Begin VSPrinter8LibCtl.VSPrinter vsPrint 
      Height          =   1815
      Left            =   600
      TabIndex        =   27
      Top             =   3720
      Visible         =   0   'False
      Width           =   5895
      _cx             =   10398
      _cy             =   3201
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   6.34469696969697
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
      AutoLinkNavigate=   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin VB.PictureBox picLiberoArt 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   6720
      Picture         =   "frmABM.frx":011A
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   26
      ToolTipText     =   "Liberar artículo para la venta"
      Top             =   840
      Width           =   210
   End
   Begin prjHiperLink.orHiperLink hliDocumento 
      Height          =   255
      Left            =   3600
      Top             =   840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColorOver   =   16711680
      Caption         =   "Ctdo B 4580"
      MouseIcon       =   "frmABM.frx":027B
      MousePointer    =   99
      BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjFindArticulo.orArticulo tArticulo 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   1200
      Width           =   5775
      _ExtentX        =   10186
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
      Appearance      =   0
   End
   Begin VB.TextBox tMemo 
      Appearance      =   0  'Flat
      Height          =   1125
      Left            =   1200
      MaxLength       =   300
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   3360
      Width           =   5775
   End
   Begin orgCalculatorFlat.orgCalculator caVarPrecio 
      Height          =   285
      Left            =   1200
      TabIndex        =   13
      Top             =   2640
      Width           =   1635
      _ExtentX        =   2884
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
      Text            =   "0.00"
   End
   Begin AACombo99.AACombo cbTipo 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   2175
      _ExtentX        =   3836
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
   Begin VB.TextBox tNSerie 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   40
      TabIndex        =   9
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox tNombre 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1560
      Width           =   5775
   End
   Begin VB.TextBox tCodigo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   500
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "print"
            Object.ToolTipText     =   "Imprimir ficha"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   5430
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":0595
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":06A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":07B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":08CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":09DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":0AEF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":0E09
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":0F1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":1235
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":154F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABM.frx":1869
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin AACombo99.AACombo cbLocal 
      Height          =   315
      Left            =   4200
      TabIndex        =   15
      Top             =   2640
      Width           =   2655
      _ExtentX        =   4683
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
   Begin AACombo99.AACombo cbEstado 
      Height          =   315
      Left            =   4200
      TabIndex        =   11
      Top             =   1920
      Width           =   2655
      _ExtentX        =   4683
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
   Begin VB.Label lbAlta 
      BackColor       =   &H00A88D7B&
      Caption         =   "Label8"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2760
      TabIndex        =   25
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label lbMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese el código del artículo específico y luego presione enter para realizar la búsqueda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   495
      Left            =   240
      TabIndex        =   24
      Top             =   4680
      Width           =   6735
   End
   Begin VB.Shape shfac 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      FillColor       =   &H00DCFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Width           =   6975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "C&omentario:"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "&Estado:"
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "&Local:"
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lbPVenta 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "15,256.00"
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   1200
      TabIndex        =   23
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Precio Vta:"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "&Variación:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lbPrecio 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "15,256.00"
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   1200
      TabIndex        =   21
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Precio:"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Tipo:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "# &Serie:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Nombre:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Artículo:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lb1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Código:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuModificar 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuOptLineSal 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVolver 
         Caption         =   "&Salir"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuVer 
      Caption         =   "Ver"
      Begin VB.Menu MnuVerGuia 
         Caption         =   "Guía de ayuda"
      End
   End
End
Attribute VB_Name = "frmABM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private paTipoCuotaContado As Long

Private Sub ValidarPermisosPrecio()
    
    caVarPrecio.Enabled = False
    Dim miConexion As New clsConexion
    'Permisos para la aplicación para el usuario logueado. (Referencia a componente aaconexion)
    caVarPrecio.Enabled = miConexion.AccesoAlMenu("ArticuloEspecifico Precio")
    Set miConexion = Nothing
    
End Sub

Private Sub loc_Print()

    If Val(tCodigo.Tag) > 0 Then
        
        'If vsPrint.PrintDialog(pdPrinterSetup) Then
            modPrintDocumento.StarDocument
            With modPrintDocumento.regPrintCnfg
                .Comentario.Dato = tMemo.Text
                .Estado.Dato = cbEstado.Text
                .ID.Dato = tCodigo.Text
                .Local.Dato = cbLocal.Text
                .Nombre.Dato = tNombre.Text
                .NombreArticulo.Dato = tArticulo.Text
                .NroSerie.Dato = tNSerie.Text
                If tNSerie.Text <> "" Then .NroSerieCB.Dato = "*" & tNSerie.Text & "*"
                .Precio.Dato = lbPrecio.Caption
                .PrecioVenta.Dato = lbPVenta.Caption
                .Tipo.Dato = cbTipo.Text
                .Variacion.Dato = caVarPrecio.Text
            End With
            vsPrint.Orientation = orLandscape
            modPrintDocumento.PrintDocument vsPrint
           
        'End If
    End If
    
End Sub

Private Sub loc_SetNroSerie()
'veo si el artículo requiere nro de serie.
Dim rsN As rdoResultset
    Set rsN = cBase.OpenResultset("Select ArtNroSerie From Articulo Where ArtID =" & tArticulo.prm_ArtID, rdOpenDynamic, rdConcurValues)
    If Not rsN.EOF Then tNSerie.Tag = IIf(rsN(0), "1", "")
    rsN.Close
End Sub

Private Sub loc_ValidarVariacion(Optional bFocus As Boolean = False)
On Error Resume Next
Dim iAux As Currency
    If IsNumeric(lbPrecio.Caption) Then iAux = CCur(lbPrecio.Caption)
    lbPVenta.Caption = Format(iAux + caVarPrecio.Text, "#,##0.00")
    If bFocus Then cbLocal.SetFocus
End Sub

Private Sub loc_ShowMsgTxt(ByVal sTxt As String)
    lbMsg.Caption = sTxt
    Status.SimpleText = sTxt
End Sub

Private Sub loc_ShowHideMsg()
    shfac.Visible = MnuVerGuia.Checked
    lbMsg.Visible = MnuVerGuia.Checked
    Status.Visible = Not MnuVerGuia.Checked
    If MnuVerGuia.Checked Then
        Me.Height = shfac.Top + shfac.Height + 780
    Else
        Me.Height = tMemo.Top + tMemo.Height + 780 + Status.Height
    End If
End Sub

Private Sub caVarPrecio_Change()
    lbPVenta.Caption = ""
End Sub

Private Sub caVarPrecio_GotFocus()
    loc_ShowMsgTxt "Ingrese la variación (positiva o negativa) a aplicar al precio final del artículo."
End Sub

Private Sub caVarPrecio_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then loc_ValidarVariacion True
End Sub

Private Sub caVarPrecio_LostFocus()
    loc_ShowMsgTxt ""
End Sub

Private Sub caVarPrecio_Validate(Cancel As Boolean)
    loc_ValidarVariacion
End Sub

Private Sub cbEstado_GotFocus()
    loc_ShowMsgTxt "Seleccione el estado actual que tiene el artículo."
End Sub

Private Sub cbEstado_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If caVarPrecio.Enabled Then
            caVarPrecio.SetFocus
        Else
            cbLocal.SetFocus
        End If
    End If
End Sub

Private Sub cbLocal_GotFocus()
    loc_ShowMsgTxt "Indique en que local se encuentra el artículo para retirarlo."
End Sub

Private Sub cbLocal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tMemo.SetFocus
End Sub

Private Sub cbLocal_LostFocus()
    loc_ShowMsgTxt ""
End Sub

Private Sub cbTipo_GotFocus()
    loc_ShowMsgTxt "Indique el tipo para identificar al artículo."
End Sub

Private Sub cbTipo_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then tArticulo.SetFocus
End Sub

Private Sub cbTipo_LostFocus()
    loc_ShowMsgTxt ""
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
'obtengo de la registry la última posición del formulario.
    ObtengoSeteoForm Me, 500, 500
    If GetSetting(App.Title, "CnfgGuia", "AA" & Me.Name & "Guia", "1") = "1" Then
        MnuVerGuia.Checked = True
    End If
    picLiberoArt.Visible = False
    loc_ShowHideMsg
    Set tArticulo.Connect = cBase
    tArticulo.DisplayCodigoArticulo = True
'Inicializo los ctrls
    loc_SetCtrl False
    loc_CleanCtrl
    
    Cons = "Select ParValor From Parametro Where ParNombre = 'TipoCuotaContado' And ParValor Is Not Null"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then paTipoCuotaContado = RsAux!ParValor
    RsAux.Close
    
    CargoCombo "Select CodID, CodTexto From Codigos Where CodCual = 128 Order by CodTexto", cbTipo
    CargoCombo "Select SucCodigo, SucAbreviacion From Sucursal Order By SucAbreviacion", cbLocal
    With cbEstado
        .AddItem "A la venta"
        .ItemData(.NewIndex) = 1
        .AddItem "Anulado"
        .ItemData(.NewIndex) = 2
    End With
    
    If paTipoCuotaContado = 0 Then MsgBox "No se cargó el parámetro Tipo de cuota Contado, no se cargará el precio del artículo.", vbExclamation, "Atención"
    
    On Error Resume Next
    ChDir App.Path
    ChDir ("..")
    ChDir (CurDir & "\REPORTES\")
    prmPathApp = CurDir & "\"
    
    InitDevicePrinter vsPrint
    Screen.MousePointer = 0
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Error al ingresar al formulario.", Err.Description
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

'Guardamos la posición del formulario.
    GuardoSeteoForm Me
'Cerramos la conexión.
    CierroConexion
'eliminamos la referencia de orcgsa.
    Set clsGeneral = Nothing
    
    modPrintDocumento.CleanArray
    SaveSetting App.Title, "CnfgGuia", "AA" & Me.Name & "Guia", IIf(MnuVerGuia.Checked, "1", "0")
    End
    Exit Sub
End Sub

Private Sub hliDocumento_Click()
    If Val(hliDocumento.Tag) > 0 Then
        If Val(Status.Tag) = 7 Then
            EjecutarApp App.Path & "\contados a domicilio.exe", "i" & Val(hliDocumento.Tag)
        Else
            EjecutarApp App.Path & IIf(Status.Tag = "1", "\Detalle de Factura.exe", "\Visualizacion de Solicitudes.exe"), Val(hliDocumento.Tag)
        End If
    End If
End Sub

Private Sub Label1_Click()
    Foco tArticulo
End Sub

Private Sub Label10_Click()
    Foco cbLocal
End Sub

Private Sub Label11_Click()
    Foco cbEstado
End Sub

Private Sub Label12_Click()
    Foco tMemo
End Sub

Private Sub Label2_Click()
    Foco tNombre
End Sub

Private Sub Label3_Click()
    Foco tNSerie
End Sub

Private Sub Label4_Click()
    Foco cbTipo
End Sub

Private Sub Label7_Click()
On Error Resume Next
    caVarPrecio.SetFocus
End Sub

Private Sub lb1_Click()
    Foco tCodigo
End Sub

Private Sub MnuCancelar_Click()
    loc_CancelarEdicion
End Sub

Private Sub MnuEliminar_Click()
    loc_Eliminar
End Sub

Private Sub MnuGrabar_Click()
    loc_Grabar
End Sub

Private Sub MnuModificar_Click()
    loc_Edicion
End Sub

Private Sub MnuNuevo_Click()
    loc_Nuevo
End Sub

Private Sub MnuVerGuia_Click()
    MnuVerGuia.Checked = Not MnuVerGuia.Checked
    loc_ShowHideMsg
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Sub loc_Nuevo()
On Error GoTo ErrAN
    Screen.MousePointer = 11
    
'Prendo Señal que es uno nuevo.
    Toolbar1.Tag = 1
    
'Limpiamos los controles para el ingreso.
    loc_CleanCtrl
    
    tCodigo.Text = "": tCodigo.Tag = ""

'Seteamos el estado de c/control para la edición.
    loc_SetCtrl True
    
'Habilito y Desabilito Botones.
    Call Botones(False, False, False, True, True, Toolbar1, Me)
    Toolbar1.Buttons("print").Enabled = False
    
    ValidarPermisosPrecio

    
'Posicionamos en el primer control para el ingreso.
    tArticulo.SetFocus
    cbTipo.SetFocus

    Screen.MousePointer = 0
    Exit Sub
    
ErrAN:
    clsGeneral.OcurrioError "Error inesperado.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub loc_Edicion()
    
'Prendo señal que es modificación.
    Toolbar1.Tag = 2
'Habilito y Desabilito Botones y controles.
    loc_SetCtrl True
    Call Botones(False, False, False, True, True, Toolbar1, Me)
    Toolbar1.Buttons("print").Enabled = False
    
    ValidarPermisosPrecio
    
'Me posiciono en el primer elemento a editar.
    Screen.MousePointer = 0
    DoEvents
    tArticulo.SetFocus
    cbTipo.SetFocus

End Sub

Private Sub loc_Grabar()
Dim sRespuesta As String

'Hacemos los controles de datos ingresados y de validación antes de grabar
    If Not fnc_ValidateSave Then Exit Sub
        
    If MsgBox("Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
        Screen.MousePointer = 11
        On Error GoTo ErrSave
        Cons = "Select * From ArticuloEspecifico Where AEsID = " & Val(tCodigo.Tag)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        'Si tengo la señal de nuevo
        If Val(Toolbar1.Tag) = 1 Then
            RsAux.AddNew
            RsAux("AEsUsuarioAlta") = paCodigoDeUsuario
            RsAux("AEsFechaAlta") = Format(Now, "yyyy/mm/dd hh:nn:ss")
        Else
            RsAux.Edit
        End If
        RsAux("AEsModificado") = Format(Now, "yyyy/mm/dd hh:nn:ss")
        RsAux("AEsArticulo") = tArticulo.prm_ArtID
        RsAux("AEsNombre") = Trim(tNombre.Text)
        RsAux("AEsTipo") = cbTipo.ItemData(cbTipo.ListIndex)
        RsAux("AEsEstado") = cbEstado.ItemData(cbEstado.ListIndex)
        If cbLocal.ListIndex > -1 Then
            RsAux("AEsLocal") = cbLocal.ItemData(cbLocal.ListIndex)
        Else
            RsAux("AEsLocal") = Null
        End If
        If Trim(tNSerie.Text) <> "" Then RsAux("AEsNroSerie") = Trim(tNSerie.Text) Else RsAux("AEsNroSerie") = Null
        If Trim(tMemo.Text) <> "" Then RsAux("AEsComentario") = Trim(tMemo.Text) Else RsAux("AEsComentario") = Null
        RsAux("AEsVariacionPrecio") = caVarPrecio.Text
        RsAux.Update
        RsAux.Close
        
        If Val(Toolbar1.Tag) = 1 Then
            Cons = "Select Max(AEsID) From ArticuloEspecifico Where AEsNombre = '" & tNombre.Text & "'"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            tCodigo.Tag = RsAux(0)
            RsAux.Close
            Toolbar1.Tag = 2
        End If
        
    'Invocamos a cancelar p/volver a estado de no edición
        loc_CancelarEdicion
    End If
    Exit Sub
    

ErrSave:
    clsGeneral.OcurrioError "No se pudo almacenar la información, reintente.", Err.Description
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub loc_Eliminar()
On Error GoTo errDel
    'Verificar si hay datos a validar.
    If MsgBox("Confirma eliminar el artículo específico?", vbQuestion + vbYesNo, "Eliminar") = vbYes Then
        Screen.MousePointer = 11
        Cons = "Select * From ArticuloEspecifico Where AEsID = " & Val(tCodigo.Tag)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            If RsAux("AEsModificado") = CDate(tNombre.Tag) And IsNull(RsAux("AEsDocumento")) Then
                RsAux.Delete
            Else
                MsgBox "Ficha modificada por otra terminal o el artículo está asociado a un documento.", vbExclamation, "Atención"
            End If
        End If
        RsAux.Close
        'Limpiamos los controles y ponemos el formulario en su nuevo estado.
        loc_CleanCtrl
        Botones True, False, False, False, False, Toolbar1, Me
        Toolbar1.Buttons("print").Enabled = False
        Screen.MousePointer = 0
    End If
    Exit Sub
errDel:
    clsGeneral.OcurrioError "No se pudo eliminar el registro.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub loc_CancelarEdicion()
    Screen.MousePointer = 11
    loc_SetCtrl False
'Si es edición cargamos los valores si no limpiamos la ficha.
    If Val(Toolbar1.Tag) = 2 Then
        'Leo y cargo valores nuevamente al mismo tiempo seteo los controles.
        loc_DBLoadData Val(tCodigo.Tag)
    Else
        'Limpio controles.
        loc_CleanCtrl
        Botones True, False, False, False, False, Toolbar1, Me
        Toolbar1.Buttons("print").Enabled = False
    End If
    'Elimino señal de edición.
    Toolbar1.Tag = 0
    tCodigo.SetFocus
    Screen.MousePointer = 0
    
End Sub

Private Sub picLiberoArt_Click()
On Error GoTo errLibArt
    If MsgBox("¿Confirma desasignar el artículo del documento?", vbQuestion + vbYesNo, "Desasignar") = vbYes Then
        cBase.Execute "UPDATE ArticuloEspecifico SET AEsTipoDocumento = Null, AEsDocumento = Null WHERE AEsID = " & Val(tCodigo.Tag)
        loc_DBLoadData Val(tCodigo.Tag)
    End If
Exit Sub
errLibArt:
    clsGeneral.OcurrioError "Error al intentar liberar el artículo.", Err.Description, "Liberar artículo"
End Sub

Private Sub tArticulo_Change()
    If tArticulo.prm_ArtID = 0 Then
        lbPrecio.Caption = ""
        lbPVenta.Caption = ""
        tNSerie.Tag = ""
    End If
End Sub

Private Sub tArticulo_GotFocus()
    loc_ShowMsgTxt "Ingrese el código o parte del nombre del artículo y presione enter para buscarlo."
End Sub

Private Sub tArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tArticulo.prm_ArtID > 0 Then
            lbPrecio.Caption = fnc_GetPrecioArticulo(tArticulo.prm_ArtID)
            loc_ValidarVariacion
            loc_SetNroSerie
            tNombre.SetFocus
        End If
    End If
End Sub

Private Sub tArticulo_LostFocus()
    loc_ShowMsgTxt ""
End Sub

Private Sub tCodigo_Change()
    If Val(tCodigo.Tag) > 0 Then
        tCodigo.Tag = "": loc_CleanCtrl
        Botones True, False, False, False, False, Toolbar1, Me
        Toolbar1.Buttons("print").Enabled = False
    End If
End Sub

Private Sub tCodigo_GotFocus()
    With tCodigo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    loc_ShowMsgTxt "Ingrese el código del artículo específico y luego presione enter para realizar la búsqueda"
End Sub

Private Sub tCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tCodigo.Text) And Val(tCodigo.Tag) = 0 Then loc_DBLoadData tCodigo.Text
    End If
End Sub

Private Sub tCodigo_LostFocus()
    loc_ShowMsgTxt ""
End Sub

Private Sub tMemo_GotFocus()
    loc_ShowMsgTxt "Ingrese un comentario para el artículo."
End Sub

Private Sub tMemo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: loc_Grabar
End Sub

Private Sub tMemo_LostFocus()
    loc_ShowMsgTxt ""
End Sub

Private Sub tNombre_GotFocus()
    loc_ShowMsgTxt "Ingrese el nombre específico para el artículo."
    With tNombre
        If Trim(.Text) = "" And Trim(tArticulo.Text) <> "" Then .Text = tArticulo.Text & " " & cbTipo.Text
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tNombre_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If tNombre.Text <> "" Then tNSerie.SetFocus
    End If
End Sub

Private Sub tNombre_LostFocus()
    loc_ShowMsgTxt ""
End Sub

Private Sub tNSerie_GotFocus()
    loc_ShowMsgTxt "Ingrese el número de serie del artículo."
    With tNSerie
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tNSerie_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then Foco cbEstado
End Sub

Private Sub tNSerie_LostFocus()
    loc_ShowMsgTxt ""
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": loc_Nuevo
        Case "modificar": loc_Edicion
        Case "eliminar": loc_Eliminar
        Case "grabar": loc_Grabar
        Case "cancelar": loc_CancelarEdicion
        Case "salir": Unload Me
        Case "print": loc_Print
    End Select

End Sub

Private Function fnc_ValidateSave() As Boolean
'Arrancamos la función en false (valor x defecto) para salir al primer error que encontremos.
    fnc_ValidateSave = False
    If tArticulo.prm_ArtID = 0 Then
        MsgBox "Debe ingresar un artículo.", vbExclamation, "Atención"
        tArticulo.SetFocus
        Exit Function
    End If
    If Trim(tNombre.Text) = "" Then
        MsgBox "Debe ingresar el nombre del artículo específico.", vbExclamation, "Atención"
        tNombre.SetFocus
        Exit Function
    End If
    If Trim(tNSerie.Text) = "" And Val(tNSerie.Tag) = 1 Then
        MsgBox "Debe ingresar el número de serie del artículo específico.", vbExclamation, "Atención"
        tNSerie.SetFocus
        Exit Function
    End If
    '201203 no dejo ingresar código de artículo como serie del mismo.
    If Trim(tNSerie.Text) = CStr(tArticulo.GetField("ArtCodigo")) Then
        MsgBox "El código del artículo no es un número de serie, corrija el dato.", vbExclamation, "ATENCIÓN"
        tNSerie.SetFocus
        Exit Function
    End If
    
    '201204 busco si tiene un código de barras único con este dato.
    Cons = "Select ACBArticulo From ArticuloCodigoBarras Where ACBArticulo = " & tArticulo.prm_ArtID & _
            " AND ACBCodigo = '" & Replace(Trim(tNSerie.Text), "'", "''") & "' AND ACBCantidad = 1"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        RsAux.Close
        MsgBox "El dato que está ingresando es el código de barras del artículo, no puede continuar.", vbExclamation, "ATENCIÓN"
        tNSerie.SetFocus
        Exit Function
    End If
    RsAux.Close
    
    If cbTipo.ListIndex = -1 Then
        MsgBox "Asigne un tipo al artículo específico.", vbExclamation, "Atención"
        cbTipo.SetFocus
        Exit Function
    End If
    If cbEstado.ListIndex = -1 Then
        MsgBox "Asigne un estado al artículo específico.", vbExclamation, "Atención"
        cbEstado.SetFocus
        Exit Function
    End If
    If cbLocal.Text <> "" And cbLocal.ListIndex = -1 Then
        MsgBox "El local ingresado no es correcto.", vbExclamation, "Atención"
        cbLocal.SetFocus
        Exit Function
    End If
    fnc_ValidateSave = True
End Function

Private Sub loc_SetCtrl(ByVal bEdit As Boolean)
'Rutina para habilitar/deshabilitar los controles

    With tCodigo
        .Enabled = Not bEdit
        .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
    End With
    With tArticulo
        .Enabled = bEdit
        .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
    End With
    With tNombre
        .Enabled = bEdit
        .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
    End With
    With tNSerie
        .Enabled = bEdit
        .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
    End With
    With cbTipo
        .Enabled = bEdit
        .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
    End With
    With cbEstado
        .Enabled = bEdit
        .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
    End With
    With cbLocal
        .Enabled = bEdit
        .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
    End With
    With caVarPrecio
        .Enabled = bEdit
        .BackColorDisplay = IIf(.Enabled, vbWindowBackground, vbButtonFace)
    End With
    With tMemo
        .Enabled = bEdit
        .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
    End With
    
    lbPrecio.BackColor = IIf(bEdit, &HC0E0FF, vbButtonFace)
    lbPVenta.BackColor = lbPrecio.BackColor
    
End Sub

Private Sub loc_CleanCtrl()
    tArticulo.Text = ""
    tNombre.Text = "": tNombre.Tag = ""
    tNSerie.Text = "": tNSerie.Tag = ""
    cbTipo.Text = ""
    cbEstado.Text = ""
    cbLocal.Text = ""
    lbPrecio.Caption = ""
    lbPVenta.Caption = ""
    caVarPrecio.Clean
    tMemo.Text = ""
    lbAlta.Visible = False: lbAlta.Caption = ""
    hliDocumento.Caption = ""
    hliDocumento.Visible = False
    picLiberoArt.Visible = False
End Sub

Private Sub loc_DBLoadData(ByVal iID As Long)
    Status.Tag = ""
    loc_CleanCtrl
    Cons = "Select ArticuloEspecifico.*, UsuIdentificacion, " & _
        " CASE AEsTipoDocumento WHEN 2 THEN 'Solicitud' WHEN 7 THEN 'Vta.Teléf.' ELSE dbo.NombreTipoDocumento(DocTipo) END AS DocNombre, " & _
        " IsNull(DocSerie, '') DocSerie, CASE WHEN AEsTipoDocumento = 1 THEN DocNumero ELSE VTeCodigo END DocNumero, " & _
        " AEsTipoDocumento DocTipo, CASE WHEN AEsTipoDocumento = 7 THEN CASE WHEN VTeAnulado IS Null THEN 0 ELSE 1 END ELSE DocAnulado END DocAnulado, ArtEsCombo " & _
        " FROM ArticuloEspecifico INNER JOIN Usuario ON AEsUsuarioAlta = UsuCodigo " & _
        " LEFT OUTER JOIN Documento ON AesDocumento = DocCodigo AND AEsTipoDocumento = 1" & _
        " LEFT OUTER JOIN VentaTelefonica ON AEsDocumento = VTeCodigo AND AEsTipoDocumento = 7" & _
        " INNER JOIN Articulo ON AEsArticulo = ArtID " & _
        " WHERE AEsID = " & iID
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'Botones toolbar si hay o no datos.
    If Not RsAux.EOF Then
        'Cargo controles
        tCodigo.Text = RsAux("AEsID")
        tCodigo.Tag = tCodigo.Text
        tArticulo.LoadArticulo RsAux("AEsArticulo")
        tNombre.Text = Trim(RsAux("AEsNombre"))
        tNombre.Tag = RsAux("AEsModificado")
        If Not IsNull(RsAux("AEsNroSerie")) Then tNSerie.Text = Trim(RsAux("AEsNroSerie"))
        BuscoCodigoEnCombo cbTipo, RsAux("AEstipo")
        BuscoCodigoEnCombo cbEstado, RsAux("AEsEstado")
        If Not IsNull(RsAux("AEsLocal")) Then BuscoCodigoEnCombo cbLocal, RsAux("AEsLocal")
        If Not IsNull(RsAux("AEsComentario")) Then tMemo.Text = Trim(RsAux("AEsComentario"))
        If Not IsNull(RsAux("AEsVariacionPrecio")) Then caVarPrecio.Text = RsAux("AEsVariacionPrecio")
        If Not RsAux("ArtEsCombo") Then
            lbPrecio.Caption = fnc_GetPrecioArticulo(tArticulo.prm_ArtID)
        Else
            lbPrecio.Caption = Format(fnc_GetPrecioCombo, "#,##0.00")
        End If
        loc_ValidarVariacion
        loc_SetNroSerie
        lbAlta.Caption = " Alta el " & Format(RsAux("AEsFechaAlta"), "dd/mm/yy hh:nn") & " por " & Trim(RsAux("UsuIdentificacion"))
        lbAlta.Visible = True
        
        If Not IsNull(RsAux("AEsDocumento")) Then
            lbAlta.Tag = "1"
            Status.Tag = RsAux("AEsTipoDocumento")
            If RsAux("AEsTipoDocumento") = 2 Then
                picLiberoArt.Visible = True
                hliDocumento.Caption = RsAux("DocNombre") & " " & RsAux("AEsDocumento")
            Else
                hliDocumento.Caption = RsAux("DocNombre") & " " & Trim(RsAux("DocSerie")) & IIf(Len(Trim(RsAux("DocSerie"))) > 0, "-", "") & RsAux("DocNumero")
                If RsAux("DocTipo") < 3 Or RsAux("DocAnulado") Or Val(Status.Tag) = 7 Then
                    picLiberoArt.Visible = True
                ElseIf Val(Status.Tag) <> 7 Then
                    picLiberoArt.Visible = fnc_DocEnNota(RsAux("AEsDocumento"))
                End If
            End If
            hliDocumento.Tag = RsAux("AEsDocumento")
            hliDocumento.Visible = True
        Else
            lbAlta.Tag = ""
        End If
        
    Else
        lbAlta.Tag = ""
        MsgBox "No se encontró el código ingresado.", vbExclamation, "Atención"
    End If
    Botones True, Not RsAux.EOF And lbAlta.Tag = "", Not RsAux.EOF And lbAlta.Tag = "", False, False, Toolbar1, Me
    Toolbar1.Buttons("print").Enabled = Not RsAux.EOF
    RsAux.Close
        
End Sub
Private Function fnc_DocEnNota(ByVal documento As Long) As Boolean
    Dim rsN As rdoResultset
    fnc_DocEnNota = False
    Cons = "SELECT NotNota FROM Nota WHERE NotFactura = " & documento
    Set rsN = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    fnc_DocEnNota = Not rsN.EOF
    rsN.Close
End Function

Private Function fnc_GetPrecioArticulo(ByVal idArticulo As Long) As String
Dim rsP As rdoResultset
    If idArticulo = 0 Then Exit Function
    Set rsP = cBase.OpenResultset("Select PViPrecio From PrecioVigente Where PViArticulo = " & idArticulo & _
                " And PViMoneda = 1 And PViHabilitado = 1 And PViTipoCuota = " & paTipoCuotaContado, rdOpenDynamic, rdConcurValues)
    If Not rsP.EOF Then fnc_GetPrecioArticulo = Format(rsP(0), "#,##0.00")
    rsP.Close
End Function

Private Function fnc_GetPrecioCombo() As String
Dim rsP As rdoResultset
    If tArticulo.prm_ArtID = 0 Then Exit Function
    
    Dim idCombo As Long
    Dim cBonifica As Currency
    Cons = "Select PreID, PreArticulo, PreImporte From Presupuesto Where PreArtCombo = " & tArticulo.prm_ArtID _
            & " And PreMoneda = 1"
    Set rsP = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsP.EOF Then
        If rsP!PreImporte <> 0 Then cBonifica = Format(rsP!PreImporte, FormatoMonedaP)
        idCombo = rsP("PreID")
    End If
    rsP.Close
    
    Dim cPrecio As Currency
    Cons = "SELECT ParCantidad, PArArticulo FROM PresupuestoArticulo WHERE PArPresupuesto = " & idCombo
    Set rsP = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsP.EOF
        cPrecio = cPrecio + (CCur(fnc_GetPrecioArticulo(rsP("PArArticulo"))) * rsP("ParCantidad"))
        rsP.MoveNext
    Loop
    rsP.Close
    fnc_GetPrecioCombo = cPrecio - cBonifica
    
End Function

