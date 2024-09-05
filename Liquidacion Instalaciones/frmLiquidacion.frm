VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{191D08B9-4E92-4372-BF17-417911F14390}#1.5#0"; "orGridPreview.ocx"
Object = "{923DD7D8-A030-4239-BCD4-51FDB459E0FE}#4.0#0"; "orgComboCalculator.ocx"
Object = "{D851F632-A4E6-4F61-863C-9480B5EC86D9}#1.2#0"; "ORGDAT~1.OCX"
Begin VB.Form frmLiquidacion 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Liquidación de Instalaciones"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLiquidacion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin orgDateCtrl.orgDate txtDesde 
      Height          =   315
      Left            =   5760
      TabIndex        =   3
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
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
      Value           =   40911
   End
   Begin orGridPreview.GridPreview gpPrint 
      Left            =   8760
      Top             =   0
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      BeginProperty FooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   7395
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picDetalle 
      Appearance      =   0  'Flat
      BackColor       =   &H00336633&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2385
      ScaleWidth      =   9345
      TabIndex        =   17
      Top             =   4200
      Width           =   9375
      Begin VB.TextBox tComentario 
         Appearance      =   0  'Flat
         Height          =   525
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   1560
         Width           =   8175
      End
      Begin VSFlex6DAOCtl.vsFlexGrid vsBoleta 
         Height          =   1335
         Left            =   4320
         TabIndex        =   13
         Top             =   120
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2355
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
         BackColorFixed  =   -2147483636
         ForeColorFixed  =   -2147483634
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   6
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
      Begin VB.TextBox tConAmp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Top             =   720
         Width           =   3255
      End
      Begin orgCalculatorFlat.orgCalculator caPesos 
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
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
         BackColorCalculator=   -2147483633
         BackColorOperator=   -2147483636
         ForeColorOperator=   -2147483634
         Text            =   "0.00"
      End
      Begin AACombo99.AACombo cLiquidar 
         Height          =   315
         Left            =   960
         TabIndex        =   8
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
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
      Begin VB.Label lQInstall 
         BackStyle       =   0  'Transparent
         Caption         =   "Instalaciones: 100"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1080
         TabIndex        =   21
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label lBoleta 
         BackStyle       =   0  'Transparent
         Caption         =   "Boleta: $15,258.96"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6600
         TabIndex        =   20
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle de la Boleta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   0
         Width           =   3135
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Co&mentario:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Concepto:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Im&porte:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Ampliación:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lTotal 
         BackColor       =   &H00FFFFFF&
         Caption         =   "  Totales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   2160
         Width           =   9375
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsInstalaciones 
      Height          =   3495
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6165
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
      BackColorFixed  =   -2147483636
      ForeColorFixed  =   -2147483634
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   4
      GridLinesFixed  =   0
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
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   16
      Top             =   7035
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "query"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "preview"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "last"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin AACombo99.AACombo cInstalador 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   120
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
   Begin MSComctlLib.ImageList ilMenu 
      Left            =   2760
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacion.frx":0442
            Key             =   "print"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacion.frx":0554
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacion.frx":0A96
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacion.frx":0EE8
            Key             =   "play"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacion.frx":0FFA
            Key             =   "save"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLiquidacion.frx":110C
            Key             =   "last"
         EndProperty
      EndProperty
   End
   Begin orgDateCtrl.orgDate txtHasta 
      Height          =   315
      Left            =   7920
      TabIndex        =   5
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
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
      Value           =   40911
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "al:"
      Height          =   255
      Left            =   7440
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Filtrar instaladas desde:"
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   7800
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Instalador:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu MnuAcceso 
      Caption         =   "Acceso"
      Visible         =   0   'False
      Begin VB.Menu MnuAccInstalacion 
         Caption         =   "Ver Instalación"
      End
      Begin VB.Menu MnuAccDetalle 
         Caption         =   "Detalle de Factura"
      End
      Begin VB.Menu MnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSeleccionar 
         Caption         =   "Seleccionar Todos"
      End
      Begin VB.Menu MnuDeseleccionar 
         Caption         =   "Deseleccionar Todos"
      End
   End
End
Attribute VB_Name = "frmLiquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bSave As Boolean
Dim lLastID As Long

Private Sub caPesos_GotFocus()
    Status.SimpleText = " Ingrese un monto."
End Sub

Private Sub caPesos_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If cLiquidar.ListIndex > -1 And cInstalador.ListIndex > -1 Then
            
            If Not Toolbar1.Buttons("save").Enabled And vsBoleta.Rows > vsBoleta.FixedRows Then
                MsgBox "No se pueden agregar detalles a una liquidación ya impresa.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
            
            With vsBoleta
                If .Rows > 1 Then
                    If .IsSubtotal(1) Then .RemoveItem 1
                End If
                .AddItem Trim(cLiquidar.Text)
                .Cell(flexcpData, .Rows - 1, 1) = 0
                .Cell(flexcpText, .Rows - 1, 1) = Trim(tConAmp.Text)
                .Cell(flexcpText, .Rows - 1, 2) = Format(caPesos.Text, "#,##0.00")
            End With
            SetTotalBoleta
            cLiquidar.Text = ""
            caPesos.Text = 0
            tConAmp.Text = ""
            cLiquidar.SetFocus
            If Not Toolbar1.Buttons("save").Enabled Then Toolbar1.Buttons("save").Enabled = True
        End If
    End If
    
End Sub

Private Sub caPesos_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub cInstalador_Change()
    frm_CleanDato
End Sub

Private Sub cInstalador_Click()
    frm_CleanDato
End Sub

Private Sub cInstalador_GotFocus()
    Status.SimpleText = " Seleccione el instalador y presione <Enter>."
End Sub

Private Sub cInstalador_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cInstalador.ListIndex > -1 Then txtDesde.SetFocus
    End If
End Sub

Private Sub cInstalador_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub cLiquidar_GotFocus()
    With cLiquidar
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = " Seleccione un concepto de liquidación."
End Sub

Private Sub cLiquidar_KeyPress(KeyAscii As Integer)
Dim iCont As Integer
    If KeyAscii = vbKeyReturn Then
        If cLiquidar.ListIndex > -1 And cInstalador.ListIndex > -1 Then
            'Válido si ya existe el item ingresado.
            With vsBoleta
                For iCont = 1 To .Rows - 1
                    If Trim(.Cell(flexcpText, iCont, 0)) = Trim(cLiquidar.Text) Then
                        If MsgBox("El concepto a ingresar ya esta en la lista." & vbCr & "¿Desea ingresar un nuevo renglón con este concepto?", vbQuestion + vbYesNo + vbDefaultButton2, "Posible Error") = vbNo Then
                            Exit Sub
                        Else
                            Exit For
                        End If
                    End If
                Next iCont
            End With
            '.........................................................
            tConAmp.SetFocus
        Else
            If tComentario.Enabled Then tComentario.SetFocus
        End If
    End If
End Sub

Private Sub cLiquidar_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub Form_Load()
On Error Resume Next

    ObtengoSeteoForm Me
    frm_CleanDato
    With Toolbar1
        Set .ImageList = ilMenu
        .Buttons("query").Image = "play"
        .Buttons("save").Image = "save"
        .Buttons("print").Image = "print"
        .Buttons("preview").Image = "preview"
        .Buttons("last").Image = "last"
    End With
    With vsBoleta
        .Rows = 1
        .Cols = 3
        .FormatString = "<Detalle|<Ampliación|>Importe"
        .ColWidth(0) = 1235
        .ColWidth(1) = 2500
        .ColWidth(2) = 750
    End With
    With vsInstalaciones
        .Rows = 1
        .Cols = 6
        .FormatString = ">Código|>Realizada|>Cumplida|>Costo|>Cobro|>Pagar|<Documento|FMod"
        .ColWidth(0) = 900
        .ColWidth(1) = 900
        .ColWidth(2) = 900
        .ColWidth(3) = 1500
        .ColWidth(4) = 1500
        .ColWidth(5) = 1500
        .ColWidth(6) = 1500
        .ColHidden(7) = True
    End With
    
    'Cargo combos.
    Cons = "Select InsCodigo, InsNombre From Instaladores Order By InsNombre"
    CargoCombo Cons, cInstalador
    
    Cons = "Select CLiCodigo, CLiDescripcion From ConceptoLiquidacion Where CLiTipoEnte = 2 Order by CLiDescripcion"
    CargoCombo Cons, cLiquidar
    
    Dim fHeader As New StdFont, ffooter As New StdFont
    With fHeader
        .Bold = True
        .Name = "Arial"
        .Size = 11
    End With
    With ffooter
        .Bold = True
        .Name = "Tahoma"
        .Size = 10
    End With
    
    With gpPrint
        .Caption = "Liquidación de Camioneros"
        .FileName = "LiquidacionCamionero"
        .Font = Font
        Set .HeaderFont = fHeader
        .Orientation = opPortrait
        .PaperSize = 1
        .PageBorder = opTopBottom
        .MarginLeft = 500
        .MarginRight = 400
    End With
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Line1.X2 = ScaleWidth - Line1.X1
    picDetalle.Move 60, ScaleHeight - (picDetalle.ScaleHeight + Toolbar1.Height + Status.Height + 40), ScaleWidth - 120
    vsInstalaciones.Move 60, Line1.Y1 + 40, ScaleWidth - 120, picDetalle.Top - Line1.Y1 - 80
    tComentario.Width = picDetalle.ScaleWidth - tComentario.Left - 30
    lTotal.Move 0, lTotal.Top, picDetalle.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    GuardoSeteoForm Me
    CierroConexion
End Sub

Private Sub Label1_Click()
    Foco cInstalador
End Sub

Private Sub Label3_Click()
    Foco tConAmp
End Sub

Private Sub Label4_Click()
    caPesos.SetFocus
End Sub

Private Sub Label5_Click()
    Foco cLiquidar
End Sub

Private Sub frm_CleanDato()
    bSave = False
    vsInstalaciones.Rows = 1
    vsBoleta.Rows = 1
    tComentario.Text = ""
    Toolbar1.Buttons("save").Enabled = False
    lBoleta.Caption = "Boleta: $": lBoleta.Tag = "0"
    lQInstall.Caption = "Instalaciones: "
End Sub

Private Sub FindInstalaciones()
On Error GoTo errFI
Dim lAux As Long
    
    frm_CleanDato
    Screen.MousePointer = 11
    Cons = "Select * From Instalacion, Documento" & _
        " Where InsInstalador = " & cInstalador.ItemData(cInstalador.ListIndex) & _
        " And InsFechaRealizada Is Not Null And InsLiquidacion Is Null And InsAnulada Is Null " & _
        " And InsDocumento = DocCodigo"
        
        
    If IsDate(txtDesde.Text) And IsDate(txtHasta.Text) Then
        Cons = Cons & " AND InsFechaRealizada BETWEEN '" & Format(txtDesde.Value, "yyyyMMdd") & "' AND '" & Format(txtHasta.Text, "yyyyMMdd 23:59:59") & "'"
    End If
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        With vsInstalaciones
            .AddItem RsAux!InsID
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!InsFechaProm, "dd/mm/yy")
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!InsFechaRealizada, "dd/mm/yy")
            If Not IsNull(RsAux!InsDebeAbonarInst) Then .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!InsDebeAbonarInst, "#,##0.00") Else .Cell(flexcpText, .Rows - 1, 4) = "0.00"
            If Not IsNull(RsAux!InsPrecioInstalacion) Then
                .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!InsPrecioInstalacion, "#,##0.00")
            Else
                'Formato Viejo.
                .Cell(flexcpText, .Rows - 1, 3) = .Cell(flexcpText, .Rows - 1, 4)
            End If
            .Cell(flexcpText, .Rows - 1, 5) = Format(.Cell(flexcpValue, .Rows - 1, 3) - .Cell(flexcpValue, .Rows - 1, 4), "#,##0.00")
            
            If RsAux!DocTipo = 1 Then
                .Cell(flexcpText, .Rows - 1, 6) = "Ctdo. "
            Else
                .Cell(flexcpText, .Rows - 1, 6) = "Créd. "
            End If
            .Cell(flexcpText, .Rows - 1, 6) = .Cell(flexcpText, .Rows - 1, 6) & Trim(RsAux!DocSerie) & "-" & Trim(RsAux!DocNumero)
            
            lAux = RsAux!DocCodigo
            .Cell(flexcpData, .Rows - 1, 6) = lAux
            .Cell(flexcpText, .Rows - 1, 7) = RsAux!InsFechaModificacion
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    If vsInstalaciones.Rows > 1 Then
        SetTotalBoleta
        Toolbar1.Buttons("save").Enabled = True
    End If
    Screen.MousePointer = 0
    Exit Sub
errFI:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar las instalaciones.", Err.Description
    vsInstalaciones.Rows = 1
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False, Optional lLiqI As Long = 0)
On Error GoTo errImprimir

'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------

    If vsInstalaciones.Rows > 1 Or vsBoleta.Rows > 1 Then
        vsInstalaciones.ExtendLastCol = False
        vsBoleta.ExtendLastCol = False
    Else
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    s_HideUnSelect True
        
    With gpPrint
        .Header = "Liquidación de Instalaciones para: " & Trim(cInstalador.Text)
        If lLiqI > 0 Then
            .LineBeforeGrid "Código de liquidación:" & lLiqI, , , True
            .LineBeforeGrid ""
        End If
        .AddGrid vsInstalaciones.hwnd
        .AddGrid vsBoleta.hwnd
        
        .LineAfterGrid ""
        .LineAfterGrid "Totales", bbold:=True
        .LineAfterGrid lQInstall.Caption & Space(10) & lBoleta.Caption
    
        If Trim(tComentario.Text) <> "" Then
            .LineAfterGrid ""
            .LineAfterGrid "Comentario: " & Trim(tComentario.Text)
        End If
        
        .LineAfterGrid ""
        
        .LineAfterGrid "Usuario: " & miConexion.UsuarioLogueado(False, True)
        
    End With
    
    If Imprimir Then
        gpPrint.GoPrint
    Else
        gpPrint.ShowPreview
    End If
    
    vsInstalaciones.ExtendLastCol = True
    vsBoleta.ExtendLastCol = True
    s_HideUnSelect False
    Screen.MousePointer = 0
    Exit Sub
errImprimir:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
    vsInstalaciones.ExtendLastCol = True
    vsBoleta.ExtendLastCol = True
    s_HideUnSelect False
End Sub
Private Sub SetTotalBoleta()
Dim iCont As Integer
Dim cValor As Currency, cPendiente As Currency
    
    'Recorro las instalaciones para poner el valor de los viáticos en el detalle de la boleta.
    cValor = 0
    lQInstall.Tag = "0"
    With vsInstalaciones
        For iCont = .FixedRows To .Rows - .FixedRows
            If .Cell(flexcpBackColor, iCont, 0) <> vbButtonFace Then
                cValor = cValor + .Cell(flexcpValue, iCont, 5)
                lQInstall.Tag = Val(lQInstall.Tag) + 1
            End If
        Next
    End With
    
    cPendiente = cValor
    lQInstall.Caption = "Instalaciones: " & Val(lQInstall.Tag)
    
    'Recorro la boleta para ver si tengo insertado el valor x los viáticos.
    With vsBoleta
        cValor = -1
        For iCont = .FixedRows To .Rows - .FixedRows
            If .Cell(flexcpData, iCont, 0) = 1 Then
                .Cell(flexcpText, iCont, 2) = Format(cPendiente, "#,##0.00")
                cValor = 0
                Exit For
            End If
        Next
        If cValor = -1 Then
            'Inserto el total de los viáticos.
            .AddItem "En Documentos", .FixedRows
            .Cell(flexcpText, 1, 2) = Format(cPendiente, "#,##0.00")
            .Cell(flexcpData, 1, 0) = 1
        End If
        cValor = 0
        For iCont = .FixedRows To .Rows - .FixedRows
            cValor = cValor + .Cell(flexcpValue, iCont, 2)
        Next
    End With
    lBoleta.Caption = "Boleta: $ " & Format(cValor, "#,##0.00")
    lBoleta.Tag = cValor
    
End Sub


Private Sub MnuAccDetalle_Click()
    If vsInstalaciones.Row >= vsInstalaciones.FixedRows Then
        EjecutarApp App.Path & "\Detalle de factura.exe", CStr(vsInstalaciones.Cell(flexcpData, vsInstalaciones.Row, 4))
    End If
End Sub

Private Sub MnuAccInstalacion_Click()
    If vsInstalaciones.Row >= vsInstalaciones.FixedRows Then
        EjecutarApp App.Path & "\Instalaciones.exe", "id:" & vsInstalaciones.Cell(flexcpValue, vsInstalaciones.Row, 0)
    End If
End Sub

Private Sub MnuDeseleccionar_Click()
Dim iQ As Integer
    For iQ = vsInstalaciones.FixedRows To vsInstalaciones.Rows - vsInstalaciones.FixedRows
        vsInstalaciones.Cell(flexcpBackColor, iQ, 0, , vsInstalaciones.Cols - 1) = vbWhite
    Next iQ
    SetTotalBoleta
End Sub

Private Sub MnuSeleccionar_Click()
Dim iQ As Integer
    
    For iQ = vsInstalaciones.FixedRows To vsInstalaciones.Rows - vsInstalaciones.FixedRows
        vsInstalaciones.Cell(flexcpBackColor, iQ, 0, , vsInstalaciones.Cols - 1) = vbButtonFace
    Next iQ
    SetTotalBoleta

End Sub

Private Sub tComentario_GotFocus()
    With tComentario
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Status.SimpleText = " Ingrese un comentario para la liquidación."
End Sub

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub tComentario_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub tConAmp_GotFocus()
    With tConAmp
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Status.SimpleText = " Ingrese una ampliación del concepto."
End Sub

Private Sub tConAmp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then caPesos.SetFocus
End Sub

Private Sub tConAmp_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "preview": AccionImprimir
        Case "print": AccionImprimir True
        Case "save": AccionGrabar
        Case "last"
            If prmPlantilla <> 0 Then EjecutarApp App.Path & "\appExploreMsg.exe ", prmPlantilla & ":I" & lLastID
    End Select
End Sub

Private Sub txtDesde_Change()
    frm_CleanDato
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Not txtDesde.HasValue Then
            If cInstalador.ListIndex > -1 Then
                FindInstalaciones
                If vsInstalaciones.Rows > 1 Then vsInstalaciones.SetFocus
            Else
                cInstalador.SetFocus
            End If
        Else
            txtHasta.SetFocus
        End If
    End If
End Sub

Private Sub txtHasta_Change()
    frm_CleanDato
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cInstalador.ListIndex > -1 And (txtDesde.Text = "" And txtHasta.Text = "") Or (txtDesde.HasValue And txtHasta.HasValue) Then
            FindInstalaciones
            If vsInstalaciones.Rows > 1 Then vsInstalaciones.SetFocus
        Else
            MsgBox "Filtros incorrectos.", vbExclamation, "ATENCIÓN"
        End If
    End If
End Sub

Private Sub vsBoleta_DblClick()
    With vsBoleta
        If .Row > 0 And Toolbar1.Buttons("save").Enabled Then
            If Val(.Cell(flexcpData, .Row, 0)) = 0 Then
                .RemoveItem .Row
            ElseIf Val(.Cell(flexcpData, .Row, 0)) = 1 Then
                MsgBox "Para eliminar el valor de los viáticos debe eliminarlos de la lista de instalaciones.", vbInformation, "ATENCIÓN"
            End If
            SetTotalBoleta
        End If
    End With
End Sub

Private Sub vsBoleta_GotFocus()
    Status.SimpleText = " Seleccione una fila del detalle a eliminar (espaciador o doble click)."
End Sub

Private Sub vsBoleta_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Shift <> 0 Then Exit Sub
    With vsBoleta
        If (KeyCode = vbKeySpace Or KeyCode = vbKeyDelete) And .Row > 0 And Toolbar1.Buttons("save").Enabled Then
            If Val(.Cell(flexcpData, .Row, 0)) = 0 Then
                .RemoveItem .Row
            ElseIf Val(.Cell(flexcpData, .Row, 0)) = 1 Then
                MsgBox "Para eliminar el valor de los viáticos debe eliminarlos de la lista de instalaciones.", vbInformation, "ATENCIÓN"
            End If
            SetTotalBoleta
        End If
    End With

End Sub

Private Sub vsBoleta_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub vsInstalaciones_DblClick()
    With vsInstalaciones
        If .Row >= .FixedRows And Toolbar1.Buttons("save").Enabled Then
            If .Cell(flexcpBackColor, .Row, 0) = vbButtonFace Then
                .Cell(flexcpBackColor, .Row, 0, , .Cols - 1) = vbWhite
            Else
                .Cell(flexcpBackColor, .Row, 0, , .Cols - 1) = vbButtonFace
            End If
            SetTotalBoleta
        End If
    End With
End Sub

Private Sub vsInstalaciones_GotFocus()
    Status.SimpleText = " Seleccione una fila a eliminar o agregar (espaciador o doble click)."
End Sub

Private Sub vsInstalaciones_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift <> 0 Then Exit Sub
    With vsInstalaciones
        If (KeyCode = vbKeySpace Or KeyCode = vbKeyDelete) And .Row > 0 And Toolbar1.Buttons("save").Enabled Then
            If .Cell(flexcpBackColor, .Row, 0) = vbButtonFace Then
                .Cell(flexcpBackColor, .Row, 0, , .Cols - 1) = vbWhite
            Else
                .Cell(flexcpBackColor, .Row, 0, , .Cols - 1) = vbButtonFace
            End If
            SetTotalBoleta
        End If
    End With
End Sub

Private Sub vsInstalaciones_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then cLiquidar.SetFocus
End Sub

Private Sub vsInstalaciones_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub vsInstalaciones_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And Shift = 0 And vsInstalaciones.Row >= vsInstalaciones.FixedRows Then
        PopupMenu MnuAcceso
    End If
End Sub

Private Sub AccionGrabar()
Dim lLiq As Long
Dim iCont As Integer
Dim dHora As Date

    If Not Toolbar1.Buttons("save").Enabled Then Exit Sub
    If MsgBox("¿Confirma grabar la liquidación?", vbQuestion + vbYesNo, "Grabar") = vbNo Then Exit Sub
    
    On Error GoTo errBT
    Screen.MousePointer = 11
    Toolbar1.Buttons("save").Enabled = False
    dHora = FechaDelServidor
    cBase.BeginTrans
    On Error GoTo errRB
   
    'Inserto en tabla LiquidacionCamiones la misma
    Cons = "Select * from Liquidacion Where LiqID = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.AddNew
    RsAux!LiqFecha = Format(dHora, "yyyy/mm/dd hh:nn")
    RsAux!LiqTipo = 2
    RsAux!LiqEnte = cInstalador.ItemData(cInstalador.ListIndex)
    RsAux.Update
    RsAux.Close

    Cons = "Select Max(LiqID) from Liquidacion Where LiqEnte = " & cInstalador.ItemData(cInstalador.ListIndex) & " And LiqTipo = 2"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly)
    If Not IsNull(RsAux(0)) Then lLiq = RsAux(0) Else lLiq = 0
    RsAux.Close
    '.....................................................................
    
    With vsInstalaciones
        For iCont = .FixedRows To .Rows - 1
            If .Cell(flexcpBackColor, iCont, 0) <> vbButtonFace Then
                Cons = "Select * From Instalacion Where InsID = " & .Cell(flexcpValue, iCont, 0)
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If RsAux.EOF Then
                    RsAux.Close
                Else
                    If CDate(.Cell(flexcpText, iCont, 7)) <> RsAux!InsFechaModificacion Then
                        RsAux.Close
                    Else
                        RsAux.Edit
                        RsAux!InsLiquidacion = lLiq
                        RsAux!InsFechaModificacion = Format(dHora, "yyyy/mm/dd hh:nn:ss")
                        RsAux.Update
                    End If
                End If
                RsAux.Close
            End If
        Next
    End With
    cBase.CommitTrans
    
    On Error Resume Next
    lLastID = lLiq
    act_SaveFileHTML lLiq
    AccionImprimir False, lLiq
    Foco cInstalador
    Screen.MousePointer = 0
    Exit Sub
    
errBT:
    Toolbar1.Buttons("save").Enabled = True
    clsGeneral.OcurrioError "Error al intentar iniciar la transacción.", Err.Description
    Screen.MousePointer = 0
    Exit Sub

errVA:
    Toolbar1.Buttons("save").Enabled = True
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Error al intentar almacenar la información, verifique si la instalación no fue modificada o eliminada por otra terminal.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
    
errRB:
    Resume errVA
    Exit Sub
End Sub

Private Sub act_SaveFileHTML(ByVal lIDLiq As Long)
Dim sFile As String
Dim iFile As Integer

    On Error GoTo errArmo
    
    sFile = "<HTML>" & vbCrLf & _
                "<HEAD>" & vbCrLf & _
                    "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html;charset=windows-1252"">" & vbCrLf & _
                    "<META NAME=""Generator"" >" & vbCrLf & _
                    "<TITLE>Liquidación de Instalación</TITLE>" & vbCrLf & _
                "</HEAD>" & vbCrLf & _
                    "<BODY>" & vbCrLf & vbCrLf & _
                    "<BR><b> Instalador: " & Trim(cInstalador.Text) & "<BR><br>" & _
                    "Código de liquidación: " & lIDLiq & "</b><br><br>" & vbCrLf

    vsInstalaciones.ExtendLastCol = False
    sFile = sFile & GetFlexGridToHTML(vsInstalaciones) & "<BR>" & vbCrLf & "<BR><b>Detalle de la boleta:</b><BR>" & vbCrLf
    vsInstalaciones.ExtendLastCol = True
    sFile = sFile & GetFlexGridToHTML(vsBoleta) & "<BR>" & vbCrLf
    sFile = sFile & "<BR><BR> Comentario: " & Trim(tComentario.Text) & "<BR><BR>Usuario: " & miConexion.UsuarioLogueado(False, True)
    sFile = sFile & _
        "<BR><b> Totales: " & "<BR><BR>" & vbCrLf & lQInstall.Caption & "&nbsp;&nbsp;&nbsp;&nbsp;" & lBoleta.Caption & "<BR><BR>" & vbCrLf & _
        "</BODY>" & vbCrLf & "</HTML>" & vbCrLf

    'ALMACENO EL ARCHIVO
    On Error GoTo errSaveLocal
    iFile = FreeFile
    Open prmPahtHTML & "LiqInstall" & Format(lIDLiq, "000000") & ".htm" For Output As iFile
    Print #iFile, sFile
    Close iFile
    Exit Sub
    
errArmo:
    MsgBox "Ocurrió el siguiente error al intentar crear el html para almacenar en el archivo.", vbCritical, "ATENCIÓN"
    Exit Sub
    
errSaveLocal:
    MsgBox "Atención ocurrió el siguiente error " & Err.Description & _
        "  al grabar el archivo html " & vbCrLf & "El mismo será almacenado en su terminal." & vbCrLf & _
        " COMUNIQUE ESTE PROBLEMA AL ADMINISTRADOR ", vbCritical, "ERROR"
    Open App.Path & "LiqInstall" & Format(lIDLiq, "000000") & ".htm" For Output As iFile
    Print #iFile, sFile
    Close iFile
    Exit Sub

End Sub

Private Sub s_HideUnSelect(ByVal bHide As Boolean)
On Error Resume Next
Dim iQ As Integer
    With vsInstalaciones
        For iQ = .FixedRows To .Rows - .FixedRows
            If Not bHide Then
                .RowHidden(iQ) = False
            Else
                If .Cell(flexcpBackColor, iQ, 0) = vbButtonFace Then .RowHidden(iQ) = True
            End If
        Next
    End With
End Sub

