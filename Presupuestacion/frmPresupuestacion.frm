VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form frmPresupuestacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuestación"
   ClientHeight    =   7920
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPresupuestacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fPresupuesto 
      Caption         =   "Presupuesto"
      ForeColor       =   &H00000080&
      Height          =   4155
      Left            =   60
      TabIndex        =   43
      Top             =   3480
      Width           =   6255
      Begin VB.TextBox tComentarioInterno 
         Appearance      =   0  'Flat
         Height          =   705
         Left            =   720
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   2560
         Width           =   5415
      End
      Begin VB.CheckBox chSinArreglo 
         Alignment       =   1  'Right Justify
         Caption         =   "Sin Arre&glo:"
         Height          =   255
         Left            =   2220
         TabIndex        =   17
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CheckBox ChReparado 
         Alignment       =   1  'Right Justify
         Caption         =   "Dar como &Reparado"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3360
         Width           =   1875
      End
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   1440
         Width           =   615
         _ExtentX        =   1085
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
      Begin VB.TextBox tUsuario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4740
         MaxLength       =   15
         TabIndex        =   19
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox tComentarioP 
         Appearance      =   0  'Flat
         Height          =   705
         Left            =   720
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   1800
         Width           =   5415
      End
      Begin VB.TextBox tCantidad 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1140
         Width           =   555
      End
      Begin VSFlex6DAOCtl.vsFlexGrid vsMotivo 
         Height          =   1095
         Left            =   2880
         TabIndex        =   8
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1931
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
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
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
      Begin VB.TextBox tPresupuesto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   30
         TabIndex        =   5
         Top             =   840
         Width           =   2715
      End
      Begin VB.TextBox tCosto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   11
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label lbComInterno 
         Caption         =   "Com &Interno:"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   2560
         Width           =   615
      End
      Begin VB.Label lReclamo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Es reclamo del Servicio 2153"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   120
         TabIndex        =   51
         Top             =   3720
         Width           =   6015
      End
      Begin VB.Label labTecnico 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5400
         TabIndex        =   48
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "&Técnico:"
         Height          =   255
         Left            =   4080
         TabIndex        =   18
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "&Cantidad:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1140
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Co&m.:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label labMotivo 
         Caption         =   "&Artículo: [F12 a Presupuesto]"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label15 
         Caption         =   "Man&o de Obra:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label LabEstadoPresupuesto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A PRESUPUESTAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   2520
         TabIndex        =   46
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label LabFAceptado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12-May 2000"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   960
         TabIndex        =   45
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label labCostoFinal 
         Caption         =   "Aceptado:"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   735
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   31
      Top             =   7665
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   8625
            TextSave        =   ""
            Key             =   "msg"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "suc"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fFicha 
      Caption         =   "Ficha"
      ForeColor       =   &H00000080&
      Height          =   3015
      Left            =   60
      TabIndex        =   26
      Top             =   420
      Width           =   6255
      Begin VB.TextBox tComentarioS 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   525
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   47
         Top             =   2400
         Width           =   6015
      End
      Begin VB.TextBox tNroSerie 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         MaxLength       =   40
         TabIndex        =   3
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox tServicio 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         MaxLength       =   6
         TabIndex        =   1
         Top             =   240
         Width           =   795
      End
      Begin VB.Label LabEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FG"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   960
         TabIndex        =   42
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label LabGarantia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12 Meses"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4680
         TabIndex        =   41
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Garantía:"
         Height          =   195
         Left            =   3780
         TabIndex        =   40
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Estado:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label LabFCompra 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12-May 2000"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4680
         TabIndex        =   38
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha Compra:"
         Height          =   195
         Left            =   3480
         TabIndex        =   37
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "&N°. Serie:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label labIngreso 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12-May 2000"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4680
         TabIndex        =   36
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Ingreso:"
         Height          =   195
         Left            =   3960
         TabIndex        =   35
         Top             =   240
         Width           =   735
      End
      Begin VB.Label labTelefono 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   34
         Top             =   960
         Width           =   5175
      End
      Begin VB.Label Label4 
         Caption         =   "Teléfonos:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   735
      End
      Begin VB.Label labEstadoServicio 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TALLER"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   1800
         TabIndex        =   32
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label Label3 
         Caption         =   "&Servicio:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   660
         Width           =   675
      End
      Begin VB.Label labIDArticulo 
         Caption         =   "Producto:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1380
         Width           =   795
      End
      Begin VB.Label LabCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "WALTER ADRIAN OCCHIUZZI MARTINEZ"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   28
         Top             =   660
         Width           =   5175
      End
      Begin VB.Label LabProducto 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(190111) REFRIGERADOR PANAVOX ALTO DE 10 PULGAS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   27
         Top             =   1380
         Width           =   5175
      End
   End
   Begin VB.PictureBox picBotones 
      Height          =   375
      Left            =   60
      ScaleHeight     =   315
      ScaleWidth      =   6195
      TabIndex        =   25
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton bEliminar 
         Height          =   310
         Left            =   780
         Picture         =   "frmPresupuestacion.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   50
         TabStop         =   0   'False
         ToolTipText     =   "Eliminar Reparación. [Ctrl+E]"
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton bHistoria 
         Height          =   310
         Left            =   5460
         Picture         =   "frmPresupuestacion.frx":040C
         Style           =   1  'Graphical
         TabIndex        =   49
         TabStop         =   0   'False
         ToolTipText     =   "Historia."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bReparado 
         Height          =   310
         Left            =   420
         Picture         =   "frmPresupuestacion.frx":0CD6
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Reparado. [Ctrl+R]"
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   1620
         Picture         =   "frmPresupuestacion.frx":1058
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Cancelar. [Ctrl+C]"
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bModificar 
         Height          =   310
         Left            =   60
         Picture         =   "frmPresupuestacion.frx":115A
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Modificar. [Ctrl+M]"
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bSalir 
         Height          =   310
         Left            =   5820
         Picture         =   "frmPresupuestacion.frx":12A4
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Salir. [Ctrl+X]"
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bGrabar 
         Height          =   310
         Left            =   1260
         Picture         =   "frmPresupuestacion.frx":13A6
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Grabar [Ctrl+G]."
         Top             =   0
         Width           =   310
      End
   End
   Begin ComctlLib.ImageList Image1 
      Left            =   6120
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPresupuestacion.frx":14A8
            Key             =   "historia"
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuReparar 
         Caption         =   "&Reparar"
         Shortcut        =   ^R
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuOpLinea 
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
      Begin VB.Menu MnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEditAuto 
         Caption         =   "Editar &automáticamente"
      End
   End
   Begin VB.Menu MnuIrA 
      Caption         =   "&Ir a"
      Begin VB.Menu MnuIrHistoria 
         Caption         =   "Historia de Servicios"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuSalDel 
         Caption         =   "Del Formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmPresupuestacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Cambios:
    '12-7 Agrego campo sin arreglo. Dejo ver servicios cumplidos no importa que sucursal sea.
    '   A los artículos de CGSA vuelvo todo a antes, osea dejo modificar si me pone que esta reparado
    '   le consulto si quiere cumplir el servicio. Sino dejo que valide Peña y lo cumple cuando me da reparar.
    ' Verifico si me dan ingreso un artículo en Deposito que no me lo agarren en gallinal y le den ingreso porque.
    ' me cagan el stock.
    
    '15-8-2000 Cdo pongo no reparado me falto poner la fecha de reparado.
    '16-8-2000 Agrande campo TalComentario a 300 caract.
    '17-8-2000 En accion modificar si no esta marcado el chreparado o chsinarreglo dejo modificarlo.
    '17-7-2001 Hacía el mov. de stock para los repuestos pero me faltaba restarle al StockTotal
    
    '9-8-2001 Al eliminar una reparación le quito el check de reparado o de sin reparar ya que vuelvo atras los mov. de repuestos.
    
Option Explicit
Public prmSucursal As String
Public prmServicio As Long
Private aTexto As String

Private Sub bCancelar_Click()
    CargoServicio True
End Sub

Private Sub bEliminar_Click()
    AccionEliminar
End Sub

Private Sub bGrabar_Click()
    AccionGrabar
End Sub

Private Sub bHistoria_Click()
    If Val(LabProducto.Tag) > 0 Then EjecutarApp pathApp & "\Historia Servicio" & " " & LabProducto.Tag
End Sub

Private Sub bModificar_Click()
    AccionModificar
End Sub

Private Sub bReparado_Click()
    
    tServicio.Text = tServicio.Tag
    CargoServicio True
    
    If bReparado.Enabled Then
        tComentarioP.Tag = "REP"
        tUsuario.Enabled = True: tUsuario.BackColor = Obligatorio
        tComentarioP.Enabled = True: tComentarioP.BackColor = vbWindowBackground
        tComentarioInterno.Enabled = True: tComentarioP.BackColor = vbWindowBackground
        tServicio.Enabled = False
        MeBotones False, False, False, True, True
        Foco tUsuario
    End If
    
End Sub

Private Sub bSalir_Click()
    Unload Me
End Sub

Private Sub cMoneda_GotFocus()
    With cMoneda
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub cMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco tCosto
End Sub

Private Sub cMoneda_LostFocus()
    cMoneda.SelStart = 0
End Sub

Private Sub ChReparado_Click()
    If ChReparado.Value = 1 Then chSinArreglo.Value = 0
End Sub

Private Sub ChReparado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If chSinArreglo.Enabled Then chSinArreglo.SetFocus Else Foco tUsuario
    End If
End Sub

Private Sub chSinArreglo_Click()
    If chSinArreglo.Value = 1 And ChReparado.Enabled Then ChReparado.Value = 0
End Sub

Private Sub chSinArreglo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tUsuario
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad
    
    picBotones.BorderStyle = vbBSNone
    bHistoria.Picture = Image1.ListImages("historia").ExtractIcon
    
    Cons = "Select MonCodigo, MonSigno From Moneda Order by MonSigno"
    CargoCombo Cons, cMoneda
    
    LimpioCamposFicha
    OcultoCamposFicha
    LimpioCamposPresupuesto
    OcultoCamposPresupuesto
    labMotivo.Caption = "Re&puesto: [F12 a Comb.Repuesto]": labMotivo.Tag = 0
    MeBotones False, False, False, False, False
    Status.Panels("suc").Text = "Sucursal: " & prmSucursal
    If LeoRegistro("AutoEdit") <> "" Then
        MnuEditAuto.Checked = True
    End If
    If prmServicio > 0 Then
        Screen.MousePointer = 11
        tServicio.Text = prmServicio
        CargoServicio False
        Screen.MousePointer = 0
    End If
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
   
    On Error Resume Next
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
    
End Sub
Private Sub InicializoGrilla()
    On Error Resume Next
    With vsMotivo
        .Editable = False
        .ExtendLastCol = True
        .Redraw = False
        .WordWrap = False
        .Rows = 1: .Cols = 1
        .FormatString = "Q|Repuestos Utilizados|Costo"
        .ColWidth(0) = 400
        .ColHidden(2) = True
        .Redraw = True
    End With
End Sub

Private Function BuscoSignoMoneda(IdMoneda As Long)
Dim RsMon As rdoResultset
    BuscoSignoMoneda = ""
    Cons = "Select * From Moneda Where MonCodigo = " & IdMoneda
    Set RsMon = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsMon.EOF Then BuscoSignoMoneda = Trim(RsMon!MonSigno)
    RsMon.Close
End Function

Private Sub Label15_Click()
    Foco cMoneda
End Sub
Private Sub Label17_Click()
    Foco tComentarioP
End Sub
Private Sub Label18_Click()
    Foco tCantidad
End Sub
Private Sub Label3_Click()
    Foco tServicio
End Sub
Private Sub Label6_Click()
    Foco tNroSerie
End Sub
Private Sub Label8_Click()
    Foco tUsuario
End Sub

Private Sub labMotivo_Click()
    Foco tPresupuesto
End Sub

Private Sub MnuCancelar_Click()
    CargoServicio True
End Sub

Private Sub MnuEditAuto_Click()
    MnuEditAuto.Checked = Not MnuEditAuto.Checked
    If MnuEditAuto.Checked Then
        GraboRegistro "AutoEdit", "1"
    Else
        GraboRegistro "AutoEdit", ""
    End If
End Sub

Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuIrHistoria_Click()
    If Val(LabProducto.Tag) > 0 Then EjecutarApp pathApp & "\Historia Servicio" & " " & LabProducto.Tag
End Sub

Private Sub MnuModificar_Click()
    AccionModificar
End Sub

Private Sub MnuReparar_Click()
    tUsuario.Enabled = True: tUsuario.BackColor = Obligatorio
    tServicio.Enabled = False
    MeBotones False, False, False, True, True
    Foco tUsuario
End Sub

Private Sub MnuSalDel_Click()
    Unload Me
End Sub


Private Sub tCantidad_GotFocus()
    With tCantidad
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tCantidad.Text = "" And tPresupuesto.Text = "" Then Foco tCosto: Exit Sub
        If tCantidad.Text = "" Then Exit Sub
        If Not IsNumeric(tCantidad.Text) Then MsgBox "El formato debe ser numérico.", vbInformation, "ATENCIÓN": Exit Sub
        If Val(tCantidad.Text) <= 0 Then MsgBox "Debe ingresar un valor positivo.", vbInformation, "ATENCIÓN": Exit Sub
        AgregoMotivo CLng(tPresupuesto.Tag)
        tPresupuesto.Text = "": tCantidad.Text = ""
    End If
End Sub

Private Sub tComentarioInterno_GotFocus()
    With tComentarioInterno
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tComentarioInterno_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Shift = vbCtrlMask And KeyCode = vbKeyReturn Then
        KeyCode = 0
        If ChReparado.Enabled Then ChReparado.SetFocus Else Foco tUsuario
        'SendKeys "{tab}"
    End If
End Sub

Private Sub tComentarioP_GotFocus()
    With tComentarioP
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tComentarioP_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyReturn Then tComentarioInterno.SetFocus
End Sub

Private Sub tCosto_GotFocus()
    With tCosto
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub tCosto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tComentarioP
End Sub
Private Sub tCosto_LostFocus()
    If IsNumeric(tCosto.Text) Then tCosto.Text = Format(tCosto.Text, FormatoMonedaP) Else tCosto.Text = ""
End Sub

Private Sub tNroSerie_GotFocus()
    With tNroSerie
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tNroSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(tNroSerie.Text) <> Trim(tNroSerie.Tag) Then
            'Modifico la tabla producto.
            Cons = "Select * From Producto Where ProCodigo = " & Val(LabProducto.Tag)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                If RsAux!ProFModificacion = CDate(LabFCompra.Tag) Then
                    FechaDelServidor
                    RsAux.Edit
                    RsAux!ProFModificacion = Format(gFechaServidor, sqlFormatoFH)
                    RsAux!ProNroSerie = tNroSerie.Text
                    RsAux.Update
                    RsAux.Close
                    tNroSerie.Tag = tNroSerie.Text
                    LabFCompra.Tag = gFechaServidor
                    tPresupuesto.SetFocus
                Else
                    RsAux.Close
                    MsgBox "Otra terminal modificó la ficha del producto, verifique.", vbExclamation, "ATENCIÓN"
                End If
            Else
                RsAux.Close
                MsgBox "Verifique si el producto no fue eliminado.", vbExclamation, "ATENCIÓN"
            End If
        End If
    End If
End Sub

Private Sub tPresupuesto_Change()
    tPresupuesto.Tag = ""
End Sub

Private Sub tPresupuesto_GotFocus()
    With tPresupuesto
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tPresupuesto_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF12
            If labMotivo.Tag = 0 Then
                labMotivo.Caption = "Comb. Re&puesto: [F12 a Repuesto]": labMotivo.Tag = 1
                tCantidad.Enabled = False: tCantidad.BackColor = Inactivo
            Else
                labMotivo.Caption = "Re&puesto: [F12 a Comb.Repuesto]": labMotivo.Tag = 0
                tCantidad.Enabled = True: tCantidad.BackColor = Blanco
            End If
        Case vbKeyReturn
            If tPresupuesto.Text <> "" Then
                If IsNumeric(tPresupuesto.Text) Then
                    If labMotivo.Tag = 1 Then   'Presupuesto
                        BuscoPresupuestoXCodigo tPresupuesto.Text
                    Else
                        BuscoArticuloXCodigo tPresupuesto.Text
                    End If
                Else
                    If labMotivo.Tag = 1 Then   'Presupuesto
                        BuscoPresupuestoXNombre
                    Else
                        BuscoArticuloXNombre
                    End If
                End If
            Else
                tCantidad.Text = ""
                Foco cMoneda
            End If
    End Select
End Sub

Private Sub tServicio_Change()
    tServicio.Tag = ""
    LimpioCamposFicha
    LimpioCamposPresupuesto
    OcultoCamposFicha
    OcultoCamposPresupuesto
    MeBotones False, False, False, False, False
    tServicio.Enabled = True
End Sub

Private Sub tServicio_GotFocus()
    With tServicio
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tServicio_KeyPress(KeyAscii As Integer)
On Error GoTo ErrCS

    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tServicio.Text) Then
            Screen.MousePointer = 11
            CargoServicio False
            Screen.MousePointer = 0
        Else
            If Trim(tServicio.Text) <> "" Then MsgBox "Formato incorrecto.", vbExclamation, "ATENCIÓN"
        End If
    End If
    Exit Sub
ErrCS:
    clsGeneral.OcurrioError "Ocurrio un error al cargar los datos del servicio.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub CargoServicio(bCancel As Boolean)
Dim LocReparacion As Integer

    LimpioCamposFicha
    LimpioCamposPresupuesto
    OcultoCamposFicha
    OcultoCamposPresupuesto
    tServicio.Enabled = True
    MeBotones False, False, False, False, False
    Cons = "Select * From Servicio, Producto, Articulo " _
        & " Where SerCodigo = " & Val(tServicio.Text) _
        & " And SerProducto = ProCodigo And ProArticulo = ArtID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        Screen.MousePointer = 0
        MsgBox "No existe un servicio con ese código.", vbInformation, "ATENCIÓN"
    Else
        If Not IsNull(RsAux!SerLocalReparacion) Then LocReparacion = RsAux!SerLocalReparacion Else LocReparacion = -1
        CargoDatosServicio
        RsAux.Close
        'VEO si le dejo entrar
        If Val(labEstadoServicio.Tag) = EstadoS.Taller Then
            If LocReparacion = -1 Then MsgBox "Este servicio no tiene asignado un local de reparación, verifique.", vbExclamation, "ATENCIÓN"
            CargoDatosTaller tServicio.Tag, LocReparacion
        ElseIf Val(labEstadoServicio.Tag) = EstadoS.Cumplido Then
            CargoDatosTaller tServicio.Tag, LocReparacion
            MeBotones False, False, False, False, False
        Else
            If Val(labEstadoServicio.Tag) = EstadoS.Retiro Then
                If LocReparacion = -1 Then MsgBox "Este servicio no tiene asignado un local de reparación, verifique.", vbExclamation, "ATENCIÓN"
                CargoDatosRetiro tServicio.Tag, LocReparacion
            Else
                OcultoCamposFicha
                OcultoCamposPresupuesto
            End If
        End If
    End If
    If MnuModificar.Enabled And MnuEditAuto.Checked And Not bCancel Then PasoModificar
    
End Sub
Private Sub CargoDatosServicio()
    tServicio.Tag = RsAux!SerCodigo
    
    labEstadoServicio.Caption = EstadoServicio(RsAux!SerEstadoServicio)
    labEstadoServicio.Tag = RsAux!SerEstadoServicio
    
    labIngreso.Caption = Format(RsAux!SerFecha, FormatoFP)
    labIngreso.Tag = RsAux!SerModificacion
    labIDArticulo.Tag = RsAux!ArtID
    
    If Not IsNull(RsAux!SerCostoFinal) Then LabEstadoPresupuesto.Tag = RsAux!SerCostoFinal Else LabEstadoPresupuesto.Tag = ""
    CargoDatosCliente RsAux!ProCliente
    LabProducto.Caption = "(" & Format(RsAux!ProCodigo, "#,000") & ") " & Trim(RsAux!ArtNombre)
    LabProducto.Tag = RsAux!ProCodigo
    LabFCompra.Tag = RsAux!ProFModificacion
    If Not IsNull(RsAux!ProNroSerie) Then tNroSerie.Text = Trim(RsAux!ProNroSerie): tNroSerie.Tag = Trim(RsAux!ProNroSerie)
    If Not IsNull(RsAux!ProCompra) Then LabFCompra.Caption = Format(RsAux!ProCompra, "dd/mm/yyyy")
    
    LabEstado.Caption = EstadoProducto(RsAux!SerEstadoProducto, True)
    LabEstado.Tag = RsAux!SerEstadoProducto
    If Not IsNull(RsAux!SerReclamoDe) Then lReclamo.Caption = "Es reclamo del Servicio " & RsAux!SerReclamoDe Else lReclamo.Caption = ""
    
    LabGarantia.Caption = RetornoGarantia(RsAux!ArtID)
    
    If Not IsNull(RsAux!SerComInterno) Then tComentarioInterno.Text = Trim(RsAux!SerComInterno)
    
    If Not IsNull(RsAux!SerComentario) Then tComentarioS.Text = Trim(RsAux!SerComentario)
    Dim RsSR As rdoResultset
    Cons = "Select * From ServicioRenglon, MotivoServicio " _
        & " Where SReServicio = " & RsAux!SerCodigo _
        & " And SReTipoRenglon = " & TipoRenglonS.Llamado _
        & " And SReMotivo = MSeID"
    Set RsSR = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsSR.EOF Then
        If Trim(tComentarioS.Text) <> "" Then tComentarioS.Text = Trim(tComentarioS.Text) & Chr(13) & Chr(10)
        tComentarioS.Text = Trim(tComentarioS.Text) & Trim(RsSR!MSeNombre): RsSR.MoveNext
    End If
    Do While Not RsSR.EOF
        tComentarioS.Text = Trim(tComentarioS.Text) & ", " & Trim(RsSR!MSeNombre)
        RsSR.MoveNext
    Loop
    RsSR.Close
End Sub
Private Sub LimpioCamposFicha()
    tServicio.Tag = ""
    labEstadoServicio.Caption = "": labEstadoServicio.Tag = ""
    labIngreso.Caption = "": labIngreso.Tag = ""
    LabCliente.Caption = ""
    labTelefono.Caption = ""
    LabProducto.Caption = "": LabProducto.Tag = ""
    tNroSerie.Text = "": tNroSerie.Tag = ""
    LabFCompra.Caption = "": LabFCompra.Tag = ""
    LabEstado.Caption = ""
    LabGarantia.Caption = ""
    tComentarioS.Text = ""
End Sub
Private Sub LimpioCamposPresupuesto()
    InicializoGrilla
    LabFAceptado.Caption = "": LabFAceptado.Tag = ""
    LabEstadoPresupuesto.BackColor = Azul
    LabEstadoPresupuesto.Caption = "": LabEstadoPresupuesto.Tag = ""
    tPresupuesto.Text = ""
    tCantidad.Text = ""
    cMoneda.Text = ""
    tCosto.Text = ""
    tComentarioP.Text = "": tComentarioP.Tag = ""
    tComentarioInterno.Text = "": tComentarioInterno.Tag = ""
    labTecnico.Caption = ""
    tUsuario.Text = "": tUsuario.Tag = 0
    ChReparado.Value = 0: ChReparado.Tag = ""
    chSinArreglo.Value = 0
    lReclamo.Caption = ""
End Sub
Private Sub CargoDatosCliente(idCliente As Long)
Dim RsCli As rdoResultset

    Cons = "Select * from Cliente " _
                & " Left Outer Join CPersona ON CliCodigo = CPeCliente " _
                & " Left Outer Join CEmpresa ON CliCodigo = CEmCliente " _
           & " Where CliCodigo = " & idCliente
           
    Set RsCli = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    If Not RsCli.EOF Then       'CI o RUC
        Select Case RsCli!CliTipo
        
            Case TipoCliente.Cliente
                If Not IsNull(RsCli!CliCiRuc) Then LabCliente.Caption = "(" & clsGeneral.RetornoFormatoCedula(RsCli!CliCiRuc) & ")"
                LabCliente.Caption = LabCliente.Caption & " " & Trim(Trim(Format(RsCli!CPeNombre1, "#")) & " " & Trim(Format(RsCli!CPeNombre2, "#"))) & ", " & Trim(Trim(Format(RsCli!CPeApellido1, "#")) & " " & Trim(Format(RsCli!CPeApellido2, "#")))
            Case TipoCliente.Empresa
                If Not IsNull(RsCli!CliCiRuc) Then LabCliente.Caption = "(" & Trim(RsCli!CliCiRuc) & ")"
                If Not IsNull(RsCli!CEmNombre) Then LabCliente.Caption = LabCliente.Caption & " " & Trim(RsCli!CEmFantasia)
                If Not IsNull(RsCli!CEmFantasia) Then LabCliente.Caption = LabCliente.Caption & " (" & Trim(RsCli!CEmFantasia) & ")"
        End Select
        LabCliente.Tag = RsCli!CliCodigo
    End If
    RsCli.Close
    labTelefono.Caption = TelefonoATexto(idCliente)     'Telefonos
End Sub
Private Sub OcultoCamposFicha()
    tNroSerie.Enabled = False: tNroSerie.BackColor = Inactivo
End Sub
Private Sub OcultoCamposPresupuesto()
    vsMotivo.BackColor = Inactivo
    tPresupuesto.Enabled = False: tPresupuesto.BackColor = Inactivo
    tCantidad.Enabled = False: tCantidad.BackColor = Inactivo
    cMoneda.Enabled = False: cMoneda.BackColor = Inactivo
    tCosto.Enabled = False: tCosto.BackColor = Inactivo
    tComentarioP.Enabled = False: tComentarioP.BackColor = Inactivo
    tComentarioInterno.Enabled = False: tComentarioInterno.BackColor = Inactivo
    tUsuario.Enabled = False: tUsuario.BackColor = Inactivo
    ChReparado.Enabled = False
    chSinArreglo.Enabled = False
End Sub
Private Sub MuestroCamposFicha()
    tNroSerie.Enabled = True: tNroSerie.BackColor = Blanco
End Sub
Private Sub MuestroCamposPresupuesto()
    vsMotivo.BackColor = Blanco
    tPresupuesto.Enabled = True: tPresupuesto.BackColor = Blanco
    tCantidad.Enabled = True: tCantidad.BackColor = Blanco
    cMoneda.Enabled = True: cMoneda.BackColor = Blanco
    tCosto.Enabled = True: tCosto.BackColor = Blanco
    tComentarioP.Enabled = True: tComentarioP.BackColor = Blanco
    tComentarioInterno.Enabled = True: tComentarioInterno.BackColor = Blanco
    tUsuario.Enabled = True: tUsuario.BackColor = Obligatorio
    If labMotivo.Tag = "1" Then tCantidad.Enabled = False: tCantidad.BackColor = Inactivo
    ChReparado.Enabled = True
    chSinArreglo.Enabled = True
End Sub
Private Sub CargoDatosRetiro(IdServicio As Long, IdLocalRepara As Integer)
    
    Cons = "Select * From ServicioVisita Where VisServicio = " & IdServicio _
        & " And VisTipo = " & TipoServicio.Retiro & " And VisSinEfecto = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "Atención hay errores en la base de datos, comuniquese con el administrador.", vbExclamation, "ATENCIÓN"
    Else
        If IsNull(RsAux!VisFImpresion) Then
            RsAux.Close
            MsgBox "Este retiro aún no fue impreso.", vbExclamation, "ATENCIÓN"
            OcultoCamposFicha
            OcultoCamposPresupuesto
        Else
            RsAux.Close
            If MsgBox("Este retiro no tiene ingreso al taller. ¿Desea darle el ingreso?", vbQuestion + vbYesNo, "INGRESO DE RETIRO") = vbYes Then
                If IdLocalRepara = paCodigoDeSucursal Then
                    IngresoATallerRetiro IdServicio
                Else
                    If MsgBox("El taller destino no es este. ¿Desea dar el ingreso a su taller?", vbQuestion + vbYesNo, "INGRESO DE RETIRO") = vbYes Then
                        IngresoATallerRetiro IdServicio
                    End If
                End If
            End If
        End If
    End If
End Sub
Private Sub IngresoATallerRetiro(IdServicio As Long)
Dim Usuario As String, Msg As String
    
    Usuario = "": Msg = ""
    Usuario = InputBox("Ingrese su digito de usuario.", "Grabar Traslado")
    If Not IsNumeric(Usuario) Then Exit Sub
    Usuario = BuscoUsuarioDigito(CLng(Usuario), True)
    If Val(Usuario) = 0 Then MsgBox "Usuario incorrecto.", vbExclamation, "ATENCIÓN": Exit Sub
    
    On Error GoTo ErrBT
    Screen.MousePointer = 11
    FechaDelServidor
    cBase.BeginTrans
    On Error GoTo ErrRB

    'TABLA SERVICIO
    Cons = "Select * From Servicio Where SerCodigo = " & IdServicio
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        Msg = "Otra terminal elimino el servicio"
        RsAux.Close: RsAux.Edit 'Provoco error.
    Else
        If RsAux!SerModificacion = CDate(labIngreso.Tag) Then
            RsAux.Edit
            RsAux!SerModificacion = Format(gFechaServidor, sqlFormatoFH)
            RsAux!SerEstadoServicio = EstadoS.Taller
            RsAux!SerLocalReparacion = paCodigoDeSucursal
            RsAux.Update
            RsAux.Close
        Else
            Msg = "Otra terminal modifico el servicio"
            RsAux.Close: RsAux.Edit 'Provoco error.
        End If
    End If
    
    Cons = "Insert Into Taller(TalServicio, TalFIngresoRealizado, TalFIngresoRecepcion,TalModificacion, TalUsuario) Values (" _
        & IdServicio & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', '" & Format(gFechaServidor, sqlFormatoFH) & "'" _
        & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', " & Usuario & ")"
    cBase.Execute (Cons)
    
    cBase.CommitTrans
    
    'Ajusto datos del servicio
    labEstadoServicio.Caption = EstadoServicio(EstadoS.Taller)
    labEstadoServicio.Tag = EstadoS.Taller
    labIngreso.Tag = gFechaServidor

    'Ahora cargo los datos de taller
    CargoDatosTaller IdServicio, CInt(paCodigoDeSucursal)
    
    Exit Sub
ErrBT:
    clsGeneral.OcurrioError "Ocurrio un error al intentar iniciar la transacción.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
ErrRB:
    Resume ErrVA
ErrVA:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrio un error al intentar almacenar la información." & Chr(13) & Msg, Err.Description
    Screen.MousePointer = 0
    
End Sub
Private Sub CargoDatosTaller(IdServicio As Long, IdLocalRepara As Integer)
Dim strUsuario As String
    
    FechaDelServidor
    
    Cons = "Select * From Taller Where TalServicio = " & IdServicio
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    If RsAux.EOF Then
        
        RsAux.Close
        MsgBox "Este servicio no tiene ingreso en taller, verifique.", vbExclamation, "ATENCIÓN"
        
    Else
        'Veo si el artìculo esta en traslado si fue recepcionado.
        If Not IsNull(RsAux!TalFIngresoRecepcion) Then
            
            If IsNull(RsAux!TalFSalidaRealizado) Then
                LabFAceptado.Tag = RsAux!TalModificacion
                
                'Veo las distintas posibilidades.-------------------
                If Not IsNull(RsAux!TalFAceptacion) Then LabFAceptado.Caption = Format(RsAux!TalFAceptacion, FormatoFP)
                If Not IsNull(RsAux!TalMonedaCosto) Then BuscoCodigoEnCombo cMoneda, RsAux!TalMonedaCosto
                If Not IsNull(RsAux!TalCostoTecnico) Then tCosto.Text = Format(RsAux!TalCostoTecnico, FormatoMonedaP)
                If Not IsNull(RsAux!TalComentario) Then tComentarioP.Text = modStart.f_QuitarClavesDelComentario(Trim(RsAux!TalComentario))
                If Not IsNull(RsAux!TalTecnico) Then labTecnico.Caption = BuscoUsuario(RsAux!TalTecnico, True)
                
                If RsAux!TalSinArreglo Then chSinArreglo.Value = 1 Else chSinArreglo.Value = 0
                
                If Not IsNull(RsAux!TalFReparado) Then
                    ChReparado.Tag = RsAux!TalFReparado
                    If chSinArreglo.Value = 0 Then ChReparado.Value = 1
                End If
                CargoRenglones
                
                If IsNull(RsAux!TalFAceptacion) Then
                    
                    If RsAux!TalSinArreglo Then
                        LabEstadoPresupuesto.Caption = "Sin Arreglo"
                    Else
                        'Si tiene costo final.
                        If LabEstadoPresupuesto.Tag = "" Then
                            LabEstadoPresupuesto.Caption = "A Presupuestar"
                        Else
                            LabEstadoPresupuesto.Caption = "En Espera de Aceptación"
                        End If
                    End If
                    
                    If RsAux!TalSinArreglo Then
                        If Not IsNull(RsAux!TalFReparado) Then
                            MeBotones False, False, True, False, False
                        Else
                            MeBotones False, True, False, False, False
                        End If
                    Else
                        MeBotones True, False, False, False, False
                    End If
                    
                Else
                    
                    If Not RsAux!TalAceptado Then
                        LabEstadoPresupuesto.Caption = "No Aceptado"
                        LabEstadoPresupuesto.BackColor = Colores.Rojo
                    Else
                        If IsNull(RsAux!TalFReparado) Then
                            LabEstadoPresupuesto.Caption = "Aceptado/A Reparar"
                        Else
                            LabEstadoPresupuesto.Caption = "Reparado"
                        End If
                    End If
                    
                    If IdLocalRepara = paCodigoDeSucursal Then
                        If IsNull(RsAux!TalFReparado) Then
                            If chSinArreglo.Value Or ChReparado.Value Then
                                MeBotones False, True, False, False, False
                            Else
                                MeBotones True, True, False, False, False
                            End If
                        Else
                            MeBotones False, False, True, False, False
                        End If
                    End If
                    
                End If
                RsAux.Close
                
                'Si el local no pretenece esta terminal.
                If IdLocalRepara <> paCodigoDeSucursal And ChReparado.Tag = "" Then
                    Screen.MousePointer = 0
                    If paClienteEmpresa = Val(LabCliente.Tag) Then
                        MsgBox "Este artículo pertenece a la empresa." _
                            & Chr(13) & "Si le da ingreso a su local se haran los movimientos de stock correspondientes.", vbInformation, "ATENCIÒN"
                    End If
                    If MsgBox("Este servicio está en traslado a otro local de reparación." & Chr(13) & "¿Desea cambiar el local de reparación al suyo?", vbQuestion + vbYesNo + vbDefaultButton2, "CAMBIAR LOCAL DE REPARACIÓN") = vbNo Then MeBotones False, False, False, False, False: Exit Sub
                    Screen.MousePointer = 11
                    strUsuario = vbNullString
                    strUsuario = InputBox("Ingrese su código de usuario.", "Recepción")
                    If Trim(strUsuario) = vbNullString Then
                        MsgBox "No se almacenará la información.", vbInformation, "ATENCIÓN": Exit Sub
                    Else
                        If Not IsNumeric(strUsuario) Then MsgBox "El formato ingresado no es numérico.", vbExclamation, "ATENCIÓN": Exit Sub
                    End If
                    strUsuario = BuscoUsuarioDigito(CLng(strUsuario), True)
                    If Val(strUsuario) = 0 Then MsgBox "Usuario incorrecto.", vbExclamation, "ATENCIÓN": Exit Sub
                    
                    Cons = "Select * From Servicio Where SerCodigo = " & IdServicio
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    RsAux.Edit
                    RsAux!SerLocalReparacion = paCodigoDeSucursal
                    RsAux.Update
                    RsAux.Close
                    HagoCambioDeLocalALocal CLng(strUsuario), labIDArticulo.Tag, IdServicio, CLng(IdLocalRepara)
                    
                    Screen.MousePointer = 0
                End If
            Else
                If IsNull(RsAux!TalFSalidaRecepcion) Then
                    MsgBox "Este servicio esta en traslado al local de origen.", vbExclamation, "ATENCIÓN"
                Else
                    MsgBox "Este servicio fue trasladado al local de origen.", vbExclamation, "ATENCIÓN"
                End If
                RsAux.Close
            End If
            
        Else
        
            'El servicio esta en traslado.
            RsAux.Close
            LabEstadoPresupuesto.Caption = "En Traslado"
            
            If IdLocalRepara <> paCodigoDeSucursal Then
                If MsgBox("Este servicio está asignado a otro local de reparación." & Chr(13) & "¿Desea cambiar el local de reparación al suyo?", vbQuestion + vbYesNo + vbDefaultButton2, "CAMBIAR LOCAL DE REPARACIÓN") = vbNo Then Exit Sub
                Cons = "Select * From Servicio Where SerCodigo = " & IdServicio
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                RsAux.Edit
                RsAux!SerLocalReparacion = paCodigoDeSucursal
                RsAux.Update
                RsAux.Close
            End If
            
            If MsgBox("El servicio está en traslado. ¿Desea dar la recepción del traslado?", vbQuestion + vbYesNo + vbDefaultButton2, "ATENCIÓN") = vbYes Then
                strUsuario = vbNullString
                strUsuario = InputBox("Ingrese su código de usuario.", "Recepción")
                If Trim(strUsuario) = vbNullString Then
                    MsgBox "No se almacenará la información.", vbInformation, "ATENCIÓN": Exit Sub
                Else
                    If Not IsNumeric(strUsuario) Then MsgBox "El formato ingresado no es numérico.", vbExclamation, "ATENCIÓN": Exit Sub
                End If
                strUsuario = BuscoUsuarioDigito(CLng(strUsuario), True)
                If Val(strUsuario) = 0 Then MsgBox "Usuario incorrecto.", vbExclamation, "ATENCIÓN": Exit Sub
                GraboIngresoDeTraslado IdServicio, CInt(strUsuario)
            End If
        End If
    End If
    
End Sub
Private Sub GraboIngresoDeTraslado(IdServicio As Long, idUsuario As Long)
Dim Msg As String, idCamion As Long
Dim RsSer As rdoResultset
    
    FechaDelServidor
    On Error GoTo ErrBT
    cBase.BeginTrans
    On Error GoTo ErrRB
    
    Cons = "Select * From Servicio Where SerCodigo = " & IdServicio
    Set RsSer = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsSer.EOF Then
        Msg = "Otra terminal elimino el servicio."
        RsSer.Close: RsSer.Edit 'Provoco error.
    Else
        If RsSer!SerModificacion = CDate(labIngreso.Tag) Then
            RsSer.Edit
            RsSer!SerModificacion = Format(gFechaServidor, sqlFormatoFH)
            RsSer.Update
            RsSer.Close
        Else
            Msg = "Otra terminal modifico el servicio."
            RsSer.Close: RsSer.Edit 'Provoco error.
        End If
    End If
    
    Cons = "Select * From Taller Where TalServicio = " & IdServicio
    Set RsSer = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    idCamion = RsSer!TalIngresoCamion
    
    RsSer.Edit
    RsSer!TalFIngresoRecepcion = Format(gFechaServidor, sqlFormatoFH)
    RsSer!TalUsuario = idUsuario
    RsSer!TalModificacion = Format(gFechaServidor, sqlFormatoFH)
    RsSer.Update
    RsSer.Close
    
    'Paso la mercadería que tiene el camión al local.
    If Val(LabCliente.Tag) = paClienteEmpresa Then HagoCambioDeLocal idUsuario, labIDArticulo.Tag, IdServicio, idCamion
    
    cBase.CommitTrans
    MeBotones True, False, False, False, False
    LabEstadoPresupuesto.Caption = "A Presupuestar"
    Screen.MousePointer = 0
    Exit Sub
    
ErrBT:
    clsGeneral.OcurrioError "Ocurrio un error al intentar iniciar la transacción.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
ErrRB:
    Resume ErrVA
ErrVA:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrio un error al intentar almacenar la información." & Chr(13) & Msg, Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub MeBotones(Modif As Boolean, Repar As Boolean, Elim As Boolean, Grabo As Boolean, Cance As Boolean)
    bModificar.Enabled = Modif: MnuModificar.Enabled = Modif
    bReparado.Enabled = Repar: MnuReparar.Enabled = Repar
    bEliminar.Enabled = Elim: MnuEliminar.Enabled = Elim
    bGrabar.Enabled = Grabo: MnuGrabar.Enabled = Grabo
    bCancelar.Enabled = Cance: MnuCancelar = Cance
End Sub
Private Sub BuscoPresupuestoXCodigo(IdPresupuesto As Long)
On Error GoTo ErrBP
    Screen.MousePointer = 11
    Cons = "Select * From Presupuesto " _
        & " Where PreCodigo = " & IdPresupuesto & " And PreEsPresupuesto = 1"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No existe un presupuesto con ese código.", vbInformation, "ATENCIÓN"
    Else
        'Veo si lo ingresó
        tPresupuesto.Tag = RsAux!PreID
        RsAux.Close
        AgregoMotivo tPresupuesto.Tag, True
        tPresupuesto.Text = ""
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrBP:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el presupuesto.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub BuscoPresupuestoXNombre()
On Error GoTo ErrBP
Dim aValor As Long
    Screen.MousePointer = 11
    Cons = "Select ID = PreID, Código = PreCodigo, Nombre = PreNombre From Presupuesto " _
        & " Where PreNombre Like '" & Replace(tPresupuesto.Text, " ", "%") & "%' And PreEsPresupuesto = 1"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No existe un presupuesto con ese nombre.", vbInformation, "ATENCIÓN"
    Else
        RsAux.MoveNext
        If RsAux.EOF Then
            RsAux.MoveFirst
            aValor = RsAux!ID
            AgregoMotivo aValor, True
            tPresupuesto.Text = ""
        Else
            RsAux.Close
            Dim objLista As New clsListadeAyuda
            If objLista.ActivarAyuda(cBase, Cons, 4500, 1, "Ayuda") > 0 Then
                aValor = objLista.RetornoDatoSeleccionado(0)
            End If
            Set objLista = Nothing
            If aValor <> 0 Then AgregoMotivo aValor, True: tPresupuesto.Text = ""
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrBP:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el presupuesto por nombre.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub BuscoArticuloXCodigo(IdPresupuesto As Long)
On Error GoTo ErrBP
    Screen.MousePointer = 11
    Cons = "Select * From Articulo, ArticuloGrupo " _
        & " Where ArtCodigo = " & IdPresupuesto & " And AGrGrupo =" & paRepuesto & " And ArtID = AGrArticulo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No existe un artículo con ese código o el mismo no pertenece al grupo repuesto.", vbInformation, "ATENCIÓN"
    Else
        tPresupuesto.Text = Trim(RsAux!ArtNombre)
        tPresupuesto.Tag = RsAux!ArtID
        tCantidad.Text = "1"
        Foco tCantidad
        RsAux.Close
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrBP:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el presupuesto.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub BuscoArticuloXNombre()
On Error GoTo ErrBP
    Screen.MousePointer = 11
    Cons = "Select ArtID, Código = ArtCodigo, Nombre = ArtNombre From Articulo, ArticuloGrupo " _
        & " Where ArtNombre Like '" & Replace(tPresupuesto.Text, " ", "%") & "%'" _
        & " And AGrGrupo =" & paRepuesto & " And ArtID = AGrArticulo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No existe un artículo con ese nombre o el mismo no pertenece al grupo repuesto.", vbInformation, "ATENCIÓN"
    Else
        RsAux.MoveNext
        If RsAux.EOF Then
            RsAux.MoveFirst
            tPresupuesto.Text = Trim(RsAux!Nombre)
            tPresupuesto.Tag = RsAux!ArtID
            RsAux.Close
            tCantidad.Text = "1"
            Foco tCantidad
        Else
            RsAux.Close
            Dim objLista As New clsListadeAyuda
            Dim aValor As Long
            If objLista.ActivarAyuda(cBase, Cons, 4500, 1, "Ayuda") > 0 Then
                aValor = objLista.RetornoDatoSeleccionado(0)
            End If
            Set objLista = Nothing
            If aValor <> 0 Then
                Cons = "Select * From Articulo Where ArtID = " & aValor
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
                tPresupuesto.Text = Trim(RsAux!ArtNombre)
                tPresupuesto.Tag = RsAux!ArtID
                RsAux.Close
                tCantidad.Text = "1"
                Foco tCantidad
            End If
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrBP:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el artículo por nombre.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub AgregoMotivo(idMotivo As Long, Optional Presupuesto As Boolean = False)
    On Error GoTo errAgregar
    
    Screen.MousePointer = 11
    Dim aValor As Long, I As Integer
        
    If Presupuesto Then
        Cons = "Select * from PresupuestoArticulo, Articulo " & _
                                 "Left Outer Join PrecioVigente On ArtID = PViArticulo " & _
                                                                          " And PViTipoCuota = " & paTipoCuotaContado & _
                                                                          " And PViMoneda = " & paMonedaPesos & _
                    " Where PArPresupuesto = " & idMotivo & _
                    " And PArArticulo = ArtID"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            With vsMotivo
                If Not ArticuloIngresado(RsAux!ArtID) Then
                    .AddItem RsAux!PArCantidad
                    aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
                    .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!ArtCodigo, "(#,000,000)") & " " & Trim(RsAux!ArtNombre)
                    If Not IsNull(RsAux!PViPrecio) Then
                        .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!PViPrecio, FormatoMonedaP)
                    Else
                        .Cell(flexcpText, .Rows - 1, 2) = RetornoPrecioDolarEnPesos(RsAux!ArtID)
                    End If
                End If
            End With
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        'Agrego el Articulo del Preosupuesto (Bonificacion)
        Cons = "Select * from Presupuesto, Articulo" & _
                   " Where PreID = " & idMotivo & _
                   " And PreArticulo = ArtID"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            With vsMotivo
                If Not ArticuloIngresado(RsAux!ArtID) Then
                    .AddItem "1"
                    aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
                    .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!ArtCodigo, "(#,000,000)") & " " & Trim(RsAux!ArtNombre)
                    .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!PreImporte, FormatoMonedaP)
                End If
            End With
        End If
        RsAux.Close
        tPresupuesto.Text = "": tCantidad.Text = "": Foco tPresupuesto
    Else
        Cons = "Select * from Articulo " & _
                                 "Left Outer Join PrecioVigente On ArtID = PViArticulo " & _
                                                                          " And PViTipoCuota = " & paTipoCuotaContado & _
                                                                          " And PViMoneda = " & paMonedaPesos & _
                    " Where ArtId = " & idMotivo
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            With vsMotivo
                If Not ArticuloIngresado(RsAux!ArtID) Then
                    .AddItem tCantidad.Text
                    aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
                    .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!ArtCodigo, "(#,000,000)") & " " & Trim(RsAux!ArtNombre)
                    If Not IsNull(RsAux!PViPrecio) Then
                        .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!PViPrecio, FormatoMonedaP)
                    Else
                        .Cell(flexcpText, .Rows - 1, 2) = RetornoPrecioDolarEnPesos(RsAux!ArtID)
                    End If
                End If
            End With
        End If
        RsAux.Close
        tPresupuesto.Text = "": tCantidad.Text = "": Foco tPresupuesto
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
errAgregar:
    clsGeneral.OcurrioError "Ocurrió un error al agregar el item a la lista.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Function RetornoPrecioDolarEnPesos(IDArticulo As Long) As String
Dim RsPD As rdoResultset
Dim TC As Currency
    RetornoPrecioDolarEnPesos = "0.00"
    Cons = "Select * From PrecioVigente Where PViArticulo  = " & IDArticulo & _
            " And PViTipoCuota = " & paTipoCuotaContado & " And PViMoneda = " & paMonedaDolar
    Set RsPD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsPD.EOF Then
        'Encontre un precio en dolar ahora lo convierto a pesos.
        TC = TasadeCambio(CInt(paMonedaDolar), CInt(paMonedaPesos), gFechaServidor)
        RetornoPrecioDolarEnPesos = Format(RsPD!PViPrecio * TC, FormatoMonedaP)
    End If
    RsPD.Close
End Function

Private Function ArticuloIngresado(IDArticulo As Long) As Boolean

    On Error GoTo errFunction
    ArticuloIngresado = True
    With vsMotivo
        For I = 1 To .Rows - 1
            If .Cell(flexcpData, I, 0) = IDArticulo Then
                MsgBox "El artículo " & .Cell(flexcpText, I, 1) & " ya está ingresado en la lista." & Chr(vbKeyReturn) & "Para modifcar la cantidad elimínelo de la lista y vuelva a ingresarlo.", vbInformation, "Item Ingresado"
                Screen.MousePointer = 0: Exit Function
            End If
        Next
    End With
    '-----------------------------------------------------------------------------------------------------
    ArticuloIngresado = False
    Exit Function

errFunction:
End Function

Private Sub tUsuario_GotFocus()
    With tUsuario
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tUsuario.Text) Then
            tUsuario.Tag = 0
            tUsuario.Tag = BuscoUsuarioDigito(Val(tUsuario.Text), True)
            If Val(tUsuario.Tag) > 0 Then AccionGrabar
        Else
            tUsuario.Tag = 0
            MsgBox "Ingrese su dígito de usuario.", vbExclamation, "ATENCIÓN"
        End If
    End If
End Sub

Private Sub AccionGrabar()
    
    If tComentarioP.Tag = "" Then
        If Not ValidoDatos Then Exit Sub
        If MsgBox("¿Confirma modificar los datos del servicio?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
        If ChReparado.Value = 1 And ChReparado.Enabled And vsMotivo.Rows = 1 And Trim(tComentarioP.Text) = "" Then MsgBox "Ingrese un repuesto o un comentario.", vbExclamation, "ATENCIÓN": Exit Sub
        GraboFicha
    Else
        If Val(tUsuario.Tag) = 0 Then MsgBox "Ingrese su digito de usuario.", vbExclamation, "ATENCIÓN": Foco tUsuario: Exit Sub
        If MsgBox("¿Confirma dar por reparado el servicio?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
        AccionReparar
    End If
    
    
End Sub

Private Function ValidoDatos() As Boolean
    ValidoDatos = False
    If Val(tUsuario.Tag) = 0 Then MsgBox "Ingrese su digito de usuario.", vbExclamation, "ATENCIÓN": Foco tUsuario: Exit Function
    If Not clsGeneral.TextoValido(tComentarioP.Text) Then MsgBox "Ingreso alguna comilla simple en el comentario del servicio.", vbExclamation, "ATENCIÓN": Foco tComentarioP: Exit Function
    If tCosto.Text <> "" Then
        If Not IsNumeric(tCosto.Text) Then MsgBox "El costo debe ser numérico.", vbExclamation, "ATENCIÓN": Foco tCosto: Exit Function
        If cMoneda.ListIndex = -1 Then MsgBox "Seleccione una moneda.", vbExclamation, "ATENCIÓN": Foco cMoneda: Exit Function
    End If
    ValidoDatos = True
End Function

Private Sub AccionModificar()
    
    tServicio.Text = tServicio.Tag
    CargoServicio True
    PasoModificar
    
    
End Sub

Private Sub PasoModificar()
    
    tComentarioP.Tag = ""
    If cMoneda.ListIndex = -1 Then BuscoCodigoEnCombo cMoneda, paMonedaPesos
    If Trim(tCosto.Text) = "" Then tCosto.Text = "0.00"
    If tServicio.Text <> tServicio.Tag Then Exit Sub
    
    MeBotones False, False, False, True, True
    MuestroCamposFicha
    MuestroCamposPresupuesto
    tServicio.Enabled = False
    
    'Si tiene costo final o tiene fecha de Aceptado.
    If LabEstadoPresupuesto.Tag <> "" Or ChReparado.Tag <> "" Then
        
        tPresupuesto.Enabled = False: tPresupuesto.BackColor = Inactivo
        tCantidad.Enabled = False: tCantidad.BackColor = Inactivo
        cMoneda.BackColor = Inactivo: cMoneda.Enabled = False
        tCosto.Enabled = False: tCosto.BackColor = Inactivo
        
        If ChReparado.Value <> 0 Or chSinArreglo.Value <> 0 Then
            ChReparado.Enabled = False
            chSinArreglo.Enabled = False
        End If
        
        If LabEstadoPresupuesto.BackColor = Azul Then
            vsMotivo.BackColor = Inactivo: Foco tComentarioP
        Else
            vsMotivo.SetFocus
        End If
        
    Else
        
        'Ingresa a presupuesta.
        Foco tPresupuesto
        
    End If

End Sub
Private Sub GraboFicha()
Dim sReparado As Boolean
    
    If Val(LabCliente.Tag) = paClienteEmpresa And chSinArreglo.Value = 1 Then
        MsgBox "El servicio será dado por cumplido y el artículo cambiara a Estado ROTO en el stock.", vbInformation, "ATENCIÓN"
    ElseIf Val(LabCliente.Tag) = paClienteEmpresa And ChReparado.Value = 1 Then
        MsgBox "El servicio es de un producto que pertenece a la empresa," & Chr(13) & "por lo tanto se dará por Cumplido." & Chr(13) _
            & "Debe quitar la hoja de servicio del producto y el mismo se tratará como un " & Chr(13) & " ARTICULO EN ESTADO SANO.", vbInformation, "ATENCIÓN"
    End If
    
    sReparado = False   'Me fijo si antes estaba reparado.
    
    Screen.MousePointer = 11
    FechaDelServidor
    On Error GoTo ErrBT
    cBase.BeginTrans
    On Error GoTo ErrVA
    Cons = "Select * From Servicio Where SerCodigo = " & tServicio.Tag
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close: cBase.RollbackTrans
        MsgBox "Otra terminal pudó eliminar el servicio verifique.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0: Exit Sub
    Else
        If RsAux!SerModificacion = CDate(labIngreso.Tag) Then
            RsAux.Edit
            RsAux!SerModificacion = Format(gFechaServidor, sqlFormatoFH)
            'Si lo marca como roto o reparado cumplo el servicio.
            If Val(LabCliente.Tag) = paClienteEmpresa And (chSinArreglo.Value = 1 Or ChReparado.Value = 1) Then
                RsAux!SerEstadoServicio = EstadoS.Cumplido
                If IsNull(RsAux!SerFCumplido) Then
                    RsAux!SerFCumplido = Format(gFechaServidor, sqlFormatoFH)
                End If
            End If
            If tComentarioInterno.Text <> "" Then
                RsAux!SerComInterno = tComentarioInterno.Text
            Else
                RsAux!SerComInterno = Null
            End If
            RsAux.Update: RsAux.Close
        Else
            RsAux.Close: cBase.RollbackTrans
            MsgBox "No podrá almacenar los datos debido a que otra terminal modifico los datos, verifique.", vbExclamation, "ATENCIÓN"
            Screen.MousePointer = 0: Exit Sub
        End If
    End If
    
    Cons = "Select * From Taller Where TalServicio = " & tServicio.Tag
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close: cBase.RollbackTrans
        MsgBox "Otra terminal pudó eliminar los datos de taller.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0: Exit Sub
    Else
    
        If RsAux!TalModificacion = CDate(LabFAceptado.Tag) Then
            
            'Veo si ya fue reparado.
            If Not IsNull(RsAux!TalFReparado) Then sReparado = True
            
            RsAux.Edit
            RsAux!TalModificacion = Format(gFechaServidor, sqlFormatoFH)
            RsAux!TalFPresupuesto = Format(gFechaServidor, sqlFormatoFH)
            RsAux!TalTecnico = tUsuario.Tag
            If cMoneda.ListIndex = -1 Then
                RsAux!TalMonedaCosto = paMonedaPesos
                RsAux!TalCostoTecnico = 0
            Else
                RsAux!TalMonedaCosto = cMoneda.ItemData(cMoneda.ListIndex)
                RsAux!TalCostoTecnico = CCur(tCosto.Text)
            End If
            
            'Veo si doy como reparado el servicio.
            If ChReparado.Value = 1 Then RsAux!TalFReparado = Format(gFechaServidor, sqlFormatoFH) Else RsAux!TalFReparado = Null
            If chSinArreglo.Value = 1 Then
                RsAux!TalSinArreglo = 1
                RsAux!TalFReparado = Format(gFechaServidor, sqlFormatoFH)
'                RsAux!TalFAceptacion = Format(gFechaServidor, sqlFormatoFH)
            Else
                RsAux!TalSinArreglo = 0
            End If

            If Not IsNull(RsAux!TalComentario) Then tComentarioP.Tag = modStart.f_GetEventos(RsAux!TalComentario) Else tComentarioP.Tag = ""
            If Trim(tComentarioP.Text) = "" And Trim(tComentarioP.Tag) = "" Then
                RsAux!TalComentario = Null
            Else
                RsAux!TalComentario = Trim(tComentarioP.Tag) & Trim(tComentarioP.Text)
            End If
            RsAux.Update: RsAux.Close
        Else
            RsAux.Close: cBase.RollbackTrans
            MsgBox "No podrá almacenar los datos debido a que otra terminal modifico los datos, verifique.", vbExclamation, "ATENCIÓN"
            Screen.MousePointer = 0: Exit Sub
        End If
    End If
    
    If vsMotivo.BackColor = Blanco Or ChReparado.Enabled Or chSinArreglo.Enabled Then
        
        'Or (ChReparado.Value = 1 And Not ChReparado.Enabled And paClienteEmpresa = Val(LabCliente.Tag))
        
        If (ChReparado.Value = 1 And ChReparado.Enabled) Or (chSinArreglo.Value = 1 And chSinArreglo.Enabled) Then
            
            With vsMotivo
            
                For I = 1 To .Rows - 1
                    If CLng(.Cell(flexcpData, I, 1)) <> paTipoArticuloServicio Then
                        MarcoMovimientoStockFisico tUsuario.Tag, TipoLocal.Deposito, paCodigoDeSucursal, CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpText, I, 0)), paEstadoArticuloEntrega, -1, TipoDocumento.Servicio, tServicio.Tag
                        MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpText, I, 0)), paEstadoArticuloEntrega, -1
                        MarcoMovimientoStockTotal CLng(.Cell(flexcpData, I, 0)), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, CCur(.Cell(flexcpText, I, 0)), -1
                    End If
                Next I
                
                Cons = "Delete ServicioRenglon Where SReServicio = " & tServicio.Tag & " And SReTipoRenglon = " & TipoRenglonS.Cumplido
                cBase.Execute (Cons)
                
                For I = 1 To .Rows - 1
                    Cons = "Insert Into ServicioRenglon (SReServicio, SReTipoRenglon, SReMotivo, SReCantidad, SReTotal) Values (" _
                        & tServicio.Tag & ", " & TipoRenglonS.Cumplido & ", " & Val(.Cell(flexcpData, I, 0)) & ", " & Val(.Cell(flexcpText, I, 0)) & ", " & CCur(.Cell(flexcpText, I, 2)) & ")"
                    cBase.Execute (Cons)
                Next I
                
            End With
            
        Else
            
            If ChReparado.Tag = "" Then
                Cons = "Delete ServicioRenglon Where SReServicio = " & tServicio.Tag & " And SReTipoRenglon = " & TipoRenglonS.Cumplido
                cBase.Execute (Cons)
                With vsMotivo
                    For I = 1 To .Rows - 1
                        Cons = "Insert Into ServicioRenglon (SReServicio, SReTipoRenglon, SReMotivo, SReCantidad, SReTotal) Values (" _
                            & tServicio.Tag & ", " & TipoRenglonS.Cumplido & ", " & Val(.Cell(flexcpData, I, 0)) & ", " & Val(.Cell(flexcpText, I, 0)) & ", " & CCur(.Cell(flexcpText, I, 2)) & ")"
                        cBase.Execute (Cons)
                    Next I
                End With
            End If
        End If
    End If
    If ChReparado.Value = 1 And ChReparado.Enabled And paClienteEmpresa = Val(LabCliente.Tag) Then
        HagoCambioDeEstado labIDArticulo.Tag, tServicio.Tag, paEstadoArticuloEntrega
    ElseIf chSinArreglo.Value = 1 And chSinArreglo.Enabled And paClienteEmpresa = Val(LabCliente.Tag) Then
        HagoCambioDeEstado labIDArticulo.Tag, tServicio.Tag, paEstadoRoto
    End If
    cBase.CommitTrans
    On Error Resume Next
    Screen.MousePointer = 0
    CargoServicio True
    tServicio.SetFocus
    Exit Sub
ErrBT:
    clsGeneral.OcurrioError "Ocurrio un error al iniciar la transacción, reintente.", Err.Description
    Screen.MousePointer = 0: Exit Sub
ErrVA:
    Resume ErrRB
ErrRB:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrio un error al intentar grabar los datos.", Err.Description
    Screen.MousePointer = 0: Exit Sub
End Sub

Private Sub vsMotivo_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsMotivo.BackColor = Inactivo Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn: Foco cMoneda
        Case vbKeyDelete: If vsMotivo.Row > 0 Then vsMotivo.RemoveItem vsMotivo.Row
    End Select
End Sub

Private Sub CargoRenglones()
Dim RsSR As rdoResultset, aValor As Long
    vsMotivo.Rows = 1
    Cons = "Select * From ServicioRenglon, Articulo Where SReServicio = " & tServicio.Tag _
        & " And SReTipoRenglon = " & TipoRenglonS.Cumplido & " And SReMotivo = ArtID"
    Set RsSR = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not RsSR.EOF
        With vsMotivo
            .AddItem RsSR!SReCantidad
            aValor = RsSR!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
            aValor = RsSR!ArtTipo: .Cell(flexcpData, .Rows - 1, 1) = aValor
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsSR!ArtCodigo, "(#,000,000)") & " " & Trim(RsSR!ArtNombre)
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsSR!SReTotal, FormatoMonedaP)
        End With
        RsSR.MoveNext
    Loop
    RsSR.Close
End Sub

Private Sub AccionReparar()
Dim RsSR As rdoResultset
    Screen.MousePointer = 11
    
    FechaDelServidor
    On Error GoTo ErrBT
    cBase.BeginTrans
    On Error GoTo ErrVA
    Cons = "Select * From Servicio Where SerCodigo = " & tServicio.Tag
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close: cBase.RollbackTrans
        MsgBox "Otra terminal pudó eliminar el servicio verifique.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0: Exit Sub
    Else
        If RsAux!SerModificacion = CDate(labIngreso.Tag) Then
            'Si es un producto de CGSA, lo doy como cumplido.
            RsAux.Edit
            If Val(LabCliente.Tag) = paClienteEmpresa Then
                RsAux!SerEstadoServicio = EstadoS.Cumplido
                If IsNull(RsAux!SerFCumplido) Then
                    RsAux!SerFCumplido = Format(gFechaServidor, sqlFormatoFH)
                End If
            End If
            RsAux!SerModificacion = Format(gFechaServidor, sqlFormatoFH)
            If tComentarioInterno.Text <> "" Then
                RsAux!SerComInterno = tComentarioInterno.Text
            Else
                RsAux!SerComInterno = Null
            End If
            RsAux.Update: RsAux.Close
        Else
            RsAux.Close: cBase.RollbackTrans
            MsgBox "No podrá almacenar los datos debido a que otra terminal modifico los datos, verifique.", vbExclamation, "ATENCIÓN"
            Screen.MousePointer = 0: Exit Sub
        End If
    End If
    Cons = "Select * From Taller Where TalServicio = " & tServicio.Tag
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close: cBase.RollbackTrans
        MsgBox "Otra terminal pudó eliminar los datos de taller.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0: Exit Sub
    Else
        If RsAux!TalModificacion = CDate(LabFAceptado.Tag) Then
            'Veo si ya fue reparado.
            RsAux.Edit
            RsAux!TalModificacion = Format(gFechaServidor, sqlFormatoFH)
            RsAux!TalFReparado = Format(gFechaServidor, sqlFormatoFH)
            RsAux!TalTecnico = tUsuario.Tag
            RsAux.Update: RsAux.Close
        Else
            RsAux.Close: cBase.RollbackTrans
            MsgBox "No podrá almacenar los datos debido a que otra terminal modifico los datos, verifique.", vbExclamation, "ATENCIÓN"
            Screen.MousePointer = 0: Exit Sub
        End If
    End If
    
    If paClienteEmpresa = Val(LabCliente.Tag) Then HagoCambioDeEstado labIDArticulo.Tag, tServicio.Tag, paEstadoArticuloEntrega
    
    If ChReparado.Value = 0 And chSinArreglo.Value = 0 Then
        With vsMotivo
            For I = 1 To .Rows - 1
                If CLng(.Cell(flexcpData, I, 1)) <> paTipoArticuloServicio Then
                    MarcoMovimientoStockFisico tUsuario.Tag, TipoLocal.Deposito, paCodigoDeSucursal, CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpText, I, 0)), paEstadoArticuloEntrega, -1, TipoDocumento.Servicio, tServicio.Tag
                    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpText, I, 0)), paEstadoArticuloEntrega, -1
                    MarcoMovimientoStockTotal CLng(.Cell(flexcpData, I, 0)), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, CCur(.Cell(flexcpText, I, 0)), -1
                End If
            Next I
        End With
    End If
    cBase.CommitTrans
    On Error Resume Next
    Screen.MousePointer = 0
    CargoServicio True
    tServicio.SetFocus
    Exit Sub
ErrBT:
    clsGeneral.OcurrioError "Ocurrio un error al iniciar la transacción, reintente.", Err.Description
    Screen.MousePointer = 0: Exit Sub
ErrVA:
    Resume ErrRB
ErrRB:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrio un error al intentar grabar los datos.", Err.Description
    Screen.MousePointer = 0: Exit Sub
End Sub

Private Sub AccionEliminar()
Dim RsSR As rdoResultset

    If MsgBox("¿Confirma eliminar la reparación de este servicio?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
        
    Dim idUsuario As String
    idUsuario = InputBox("Ingrese su digito de usuario.", "Usuario")
    If Not IsNumeric(idUsuario) Then
        MsgBox "No ingresó un dígito correcto.", vbExclamation, "ATENCIÓN"
        Exit Sub
    Else
        tUsuario.Tag = ""
        tUsuario.Tag = BuscoUsuarioDigito(Val(idUsuario), True)
        If Val(tUsuario.Tag) = 0 Then
            MsgBox "No ingresó un dígito correcto.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
    End If

    Screen.MousePointer = 11
    FechaDelServidor
    On Error GoTo ErrBT
    cBase.BeginTrans
    On Error GoTo ErrVA
    Cons = "Select * From Servicio Where SerCodigo = " & tServicio.Tag
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close: cBase.RollbackTrans
        MsgBox "Otra terminal pudó eliminar el servicio verifique.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0: Exit Sub
    Else
        If RsAux!SerModificacion = CDate(labIngreso.Tag) Then
            'Si es un producto de CGSA, lo doy como cumplido.
            RsAux.Edit
            RsAux!SerEstadoServicio = EstadoS.Taller
            RsAux!SerModificacion = Format(gFechaServidor, sqlFormatoFH)
            RsAux.Update: RsAux.Close
        Else
            RsAux.Close: cBase.RollbackTrans
            MsgBox "No podrá almacenar los datos debido a que otra terminal modifico los datos, verifique.", vbExclamation, "ATENCIÓN"
            Screen.MousePointer = 0: Exit Sub
        End If
    End If
    Cons = "Select * From Taller Where TalServicio = " & tServicio.Tag
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close: cBase.RollbackTrans
        MsgBox "Otra terminal pudó eliminar los datos de taller.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0: Exit Sub
    Else
        If RsAux!TalModificacion = CDate(LabFAceptado.Tag) Then
            'Veo si ya fue reparado.
            RsAux.Edit
            RsAux!TalModificacion = Format(gFechaServidor, sqlFormatoFH)
            RsAux!TalFReparado = Null
            RsAux!TalSinArreglo = 0
            RsAux.Update: RsAux.Close
        Else
            RsAux.Close: cBase.RollbackTrans
            MsgBox "No podrá almacenar los datos debido a que otra terminal modifico los datos, verifique.", vbExclamation, "ATENCIÓN"
            Screen.MousePointer = 0: Exit Sub
        End If
    End If
    With vsMotivo
        For I = 1 To .Rows - 1
            If CLng(.Cell(flexcpData, I, 1)) <> paTipoArticuloServicio Then
                MarcoMovimientoStockFisico tUsuario.Tag, TipoLocal.Deposito, paCodigoDeSucursal, CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpText, I, 0)), paEstadoArticuloEntrega, 1, TipoDocumento.Servicio, tServicio.Tag
                MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, CLng(.Cell(flexcpData, I, 0)), CCur(.Cell(flexcpText, I, 0)), paEstadoArticuloEntrega, 1
                MarcoMovimientoStockTotal CLng(.Cell(flexcpData, I, 0)), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, CCur(.Cell(flexcpText, I, 0)), 1
            End If
        Next I
    End With
    cBase.CommitTrans
    On Error Resume Next
    Screen.MousePointer = 0
    CargoServicio True
    Exit Sub
ErrBT:
    clsGeneral.OcurrioError "Ocurrio un error al iniciar la transacción, reintente.", Err.Description
    Screen.MousePointer = 0: Exit Sub
ErrVA:
    Resume ErrRB
ErrRB:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrio un error al intentar grabar los datos.", Err.Description
    Screen.MousePointer = 0: Exit Sub
End Sub
Private Sub HagoCambioDeEstado(IDArticulo As Long, IdServicio As Long, NuevoEstado As Integer)
    
    'Cambio el estado del artículo como Sano a Recuperar.
    MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, paCodigoDeSucursal, IDArticulo, 1, paEstadoARecuperar, -1, TipoDocumento.ServicioCambioEstado, IdServicio
    MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, paCodigoDeSucursal, IDArticulo, 1, NuevoEstado, 1, TipoDocumento.ServicioCambioEstado, IdServicio
        
    MarcoMovimientoStockTotal IDArticulo, TipoEstadoMercaderia.Fisico, paEstadoARecuperar, 1, -1
    MarcoMovimientoStockTotal IDArticulo, TipoEstadoMercaderia.Fisico, NuevoEstado, 1, 1
    
    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, IDArticulo, 1, paEstadoARecuperar, -1
    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, IDArticulo, 1, NuevoEstado, 1
    
End Sub

Private Sub HagoCambioDeLocal(idUsuario As Long, IDArticulo As Long, IdServicio As Long, idCamion As Long)
    
    'Le hago un alta al local
    MarcoMovimientoStockFisico idUsuario, TipoLocal.Deposito, paCodigoDeSucursal, IDArticulo, 1, paEstadoARecuperar, 1, TipoDocumento.ServicioCambioEstado, IdServicio
    'Le bajo al camión.
    MarcoMovimientoStockFisico idUsuario, TipoLocal.Camion, idCamion, IDArticulo, 1, paEstadoARecuperar, -1, TipoDocumento.ServicioCambioEstado, IdServicio
    'Le hago un alta al local
    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, IDArticulo, 1, paEstadoARecuperar, 1
    'Le bajo al camión.
    MarcoMovimientoStockFisicoEnLocal TipoLocal.Camion, idCamion, IDArticulo, 1, paEstadoARecuperar, -1
    
End Sub

Private Sub HagoCambioDeLocalALocal(idUsuario As Long, IDArticulo As Long, IdServicio As Long, idLocal As Long)
    
    'Le hago un alta al local
    MarcoMovimientoStockFisico idUsuario, TipoLocal.Deposito, paCodigoDeSucursal, IDArticulo, 1, paEstadoARecuperar, 1, TipoDocumento.ServicioCambioEstado, IdServicio
    'Le bajo al camión.
    MarcoMovimientoStockFisico idUsuario, TipoLocal.Deposito, idLocal, IDArticulo, 1, paEstadoARecuperar, -1, TipoDocumento.ServicioCambioEstado, IdServicio
    'Le hago un alta al local
    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, IDArticulo, 1, paEstadoARecuperar, 1
    'Le bajo al camión.
    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, idLocal, IDArticulo, 1, paEstadoARecuperar, -1
    
End Sub

Private Function LeoRegistro(ByVal sPropiedad As String) As String
    LeoRegistro = GetSetting(App.Title, "Properties", sPropiedad, "")
End Function

Private Sub GraboRegistro(ByVal sPropiedad As String, ByVal sValor As String)
    SaveSetting App.Title, "Properties", sPropiedad, sValor
End Sub

