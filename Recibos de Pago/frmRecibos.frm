VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5694326E-AE1E-40BF-B7B0-0E8918015F0D}#1.1#0"; "orChequeCtrl.ocx"
Begin VB.Form frmRecibos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibos de Pago"
   ClientHeight    =   5010
   ClientLeft      =   2880
   ClientTop       =   3060
   ClientWidth     =   7950
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRecibos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7950
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir."
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   250
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   400
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pagos"
            Object.ToolTipText     =   "Con Qué paga"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "dolar"
            Object.ToolTipText     =   "Tasas de cambio"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Comprobante"
      ForeColor       =   &H00000080&
      Height          =   1215
      Left            =   60
      TabIndex        =   29
      Top             =   480
      Width           =   7815
      Begin VB.TextBox tProveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3480
         MaxLength       =   40
         TabIndex        =   5
         Top             =   540
         Width           =   4215
      End
      Begin VB.TextBox tID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   960
         MaxLength       =   12
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox tTCDolar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6000
         MaxLength       =   6
         TabIndex        =   13
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox tIOriginal 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   4300
         MaxLength       =   15
         TabIndex        =   11
         Text            =   "1,000,000.00"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox tNumero 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1305
         MaxLength       =   9
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   960
         MaxLength       =   12
         TabIndex        =   3
         Top             =   540
         Width           =   1215
      End
      Begin VB.TextBox tSerie 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   960
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "BB"
         Top             =   840
         Width           =   320
      End
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   3480
         TabIndex        =   10
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
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
         Text            =   "U$S"
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "I&d Compra:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lTC 
         BackStyle       =   0  'Transparent
         Caption         =   "21/08/2000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6840
         TabIndex        =   31
         Top             =   885
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "T/&C:"
         Height          =   255
         Left            =   5520
         TabIndex        =   12
         Top             =   885
         Width           =   375
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "&Importe:"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   885
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Proveedor:"
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Número:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Facturas que se Pagan"
      ForeColor       =   &H00000080&
      Height          =   2940
      Left            =   60
      TabIndex        =   28
      Top             =   1755
      Width           =   7815
      Begin VB.TextBox tFPaga 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3480
         MaxLength       =   15
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton bAgregar 
         Caption         =   "&Agregar"
         Height          =   315
         Left            =   4920
         TabIndex        =   19
         Top             =   225
         Width           =   855
      End
      Begin VB.TextBox tFSerie 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   960
         MaxLength       =   2
         TabIndex        =   15
         Text            =   "BB"
         Top             =   240
         Width           =   320
      End
      Begin VB.TextBox tFNumero 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1305
         MaxLength       =   9
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox tComentario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   25
         Top             =   2580
         Width           =   6495
      End
      Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
         Height          =   1245
         Left            =   120
         TabIndex        =   20
         Top             =   900
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   2196
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
         FocusRect       =   2
         HighLight       =   2
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
      Begin orChequeCtrl.orCheque orCheque 
         Height          =   315
         Left            =   3960
         TabIndex        =   23
         Top             =   2220
         Width           =   3150
         _ExtentX        =   5556
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
      Begin AACombo99.AACombo cDisponibilidad 
         Height          =   315
         Left            =   1200
         TabIndex        =   22
         Top             =   2220
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   556
         ForeColor       =   0
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
         Text            =   ""
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Se P&aga Con:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2280
         Width           =   1035
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lFSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1,000,000.00"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   960
         TabIndex        =   35
         Top             =   560
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Pa&ga:"
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lFFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12/12/2000"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5640
         TabIndex        =   34
         Top             =   560
         Width           =   975
      End
      Begin VB.Label lFImporte 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1,000,000.00"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3480
         TabIndex        =   33
         Top             =   560
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   4920
         TabIndex        =   32
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fac&tura:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe:"
         Height          =   255
         Left            =   2760
         TabIndex        =   26
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Comentario&s:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2580
         Width           =   975
      End
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   30
      Top             =   4755
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "terminal"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "bd"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5821
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7680
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecibos.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecibos.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecibos.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecibos.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecibos.frx":0752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecibos.frx":0A6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecibos.frx":0D86
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecibos.frx":10A0
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
      Begin VB.Menu MnuL1 
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
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuExit 
         Caption         =   "Del formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sNuevo As Boolean, sModificar As Boolean
Dim prmIdGasto As Long
Dim gFModificacion As Date

Dim rsCom As rdoResultset

Dim bPagoConDC As Boolean

Private Type typReg
    oFechaCompra As Date
    oIDMovimiento As Long       'old Id Movimiento de Disponibilidad (caso de modificacion)
    oTotalBruto As Currency     'Total Bruto, p/controlar cambios de importe cdo está pago
    oPesos As Currency              'Total bruto en pesos (p/ cambios de moneda)
    oUsuario As Long             'Para sucesos por modificacion al cerrar disponibilidad
    
    Disponibilidad As Long
    ImporteCompra As Currency
    ImporteDisponibilidad As Currency
    ImportePesos As Currency
    HaceSalidaCaja As Boolean
    
    cndPagoConOtras As Boolean                            'Pago con otras disponibilidades
    cndFCierreDisponibilidad As Date         'Fecha de Cierre de la Disponibilidad
    
    flgSucesoXMod As Boolean            'Si hay suceso por modificacion de datos
End Type

Dim mData As typReg

Private Sub bAgregar_Click()

    If Val(tFNumero.Tag) = 0 Then
        MsgBox "Ingrese los datos la factura a pagar con el recibo.", vbExclamation, "ATENCIÓN"
        Foco tFSerie: Exit Sub
    End If
    
    If Not IsNumeric(tFPaga.Text) Then
        MsgBox "Ingrese el importe a pagar para la factura seleccionada.", vbExclamation, "ATENCIÓN"
        Foco tFPaga: Exit Sub
    End If
    If CCur(tFPaga.Text) = 0 Then
        MsgBox "El importe a pagar para la factura seleccionada debe ser mayor a cero.", vbExclamation, "ATENCIÓN"
        Foco tFPaga: Exit Sub
    End If
    
    If HayPago + CCur(tFPaga.Text) > CCur(tIOriginal.Text) Then
        MsgBox "El importe ingresado (acumulado de pagos) es mayor que el importe original del recibo.", vbExclamation, "ATENCIÓN"
        Foco tFPaga: Exit Sub
    End If
    
    If CCur(tFPaga.Text) > CCur(lFSaldo.Caption) Then
        MsgBox "El importe ingresado (para saldar la factura) es mayor que el saldo de la factura.", vbExclamation, "ATENCIÓN"
        Foco tFPaga: Exit Sub
    End If
    
    On Error GoTo errAgregar
    'Verifico si la factura está en la lista----------------------------------------
    For I = 1 To vsLista.Rows - 1
        If vsLista.Cell(flexcpValue, I, 4) = Val(tFNumero.Tag) Then
            MsgBox "La factura seleccionada ya está ingresada. " & vbCrLf & _
                        "Verifique la lista de facturas pagas.", vbInformation, "Factura Inrgesada"
            Exit Sub
        End If
    Next
        
    'Agrego la factura a la lista de facturas pagas----------------------------------------
    Screen.MousePointer = 11
    With vsLista
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = Trim(tFSerie.Text)
        .Cell(flexcpText, .Rows - 1, 1) = Trim(tFNumero.Text)
        
        .Cell(flexcpText, .Rows - 1, 2) = lFFecha.Caption
        .Cell(flexcpText, .Rows - 1, 3) = lFImporte.Caption
        
        .Cell(flexcpText, .Rows - 1, 4) = Format(tFNumero.Tag, "#,##0")
        .Cell(flexcpText, .Rows - 1, 5) = Format(tFPaga.Text, "#,##0.00")
        .Cell(flexcpText, .Rows - 1, 6) = lFSaldo.Caption
    End With
    
    tFSerie.Text = "": tFNumero.Text = "": tFPaga.Text = ""
    If CCur(tIOriginal.Text) = HayPago Then Foco cDisponibilidad Else Foco tFSerie
    
    Screen.MousePointer = 0     '-------------------------------------------------------------------
    Exit Sub
    
errAgregar:
    clsGeneral.OcurrioError "Error al agregar la factura a la lista.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub cDisponibilidad_Change()
    If Val(cDisponibilidad.Tag) = 0 Then
        orCheque.fnc_BlankControls
        orCheque.Enabled = False: orCheque.BackColor = Colores.Inactivo
        cDisponibilidad.Tag = 1
    End If
End Sub

Private Sub cDisponibilidad_Click()
    If Val(cDisponibilidad.Tag) = 0 Then
        orCheque.fnc_BlankControls
        orCheque.Enabled = False: orCheque.BackColor = Colores.Inactivo
        cDisponibilidad.Tag = 1
    End If
End Sub

Private Sub cDisponibilidad_KeyDown(KeyCode As Integer, Shift As Integer)
 
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyDivide
                cDisponibilidad.ListIndex = 0
                KeyCode = 0
    
        Case vbKeyReturn
                If Not IsNumeric(tIOriginal.Text) Then Foco tIOriginal: Exit Sub
                If cMoneda.ListIndex = -1 Then Foco cMoneda: Exit Sub
                If cDisponibilidad.ListIndex = -1 Then Exit Sub
                cDisponibilidad.Tag = 0
                'Veo si la seleccionada es bancaria
                Dim mID As Long, mIdx As Integer
                mID = cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
                If mID > 0 Then
                    mIdx = dis_IdxArray(mID)
                    If mIdx <> -1 Then
                        If arrDisp(mIdx).Bancaria Then
                            orCheque.Enabled = True: orCheque.BackColor = vbWindowBackground
                            With orCheque
                                .prp_IdDisponibilidad = mID
                                .prp_IdGasto = prmIdGasto
                                .prp_ValorAAsignar = CCur(tIOriginal.Text)
                                .prp_IdCheque = Val(orCheque.Tag)
                                .prp_ValorInicial = 1
                                .fnc_Show
                            End With
                        End If
                    End If
                End If
                
                If orCheque.Enabled Then orCheque.SetFocus Else Foco tComentario
        
    End Select

End Sub

Private Sub cDisponibilidad_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "/" Then KeyAscii = 0
End Sub

Private Sub cMoneda_Change()
    If cDisponibilidad.ListCount > 0 Then cDisponibilidad.Clear
End Sub

Private Sub cMoneda_Click()
    If cDisponibilidad.ListCount > 0 Then cDisponibilidad.Clear
End Sub

Private Sub cMoneda_GotFocus()
    cMoneda.SelStart = 0: cMoneda.SelLength = Len(cMoneda.Text)
End Sub

Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If cMoneda.ListIndex = -1 Then Exit Sub
        
        If IsDate(tFecha.Text) Then
            Dim aFechaTC As String: aFechaTC = ""
            tTCDolar.ToolTipText = "": lTC.Caption = ""
            
            'TC del ultimo dia del mes anterior
            tTCDolar.Text = TasadeCambio(paMonedaDolar, paMonedaPesos, UltimoDia(DateAdd("m", -1, CDate(tFecha.Text))), aFechaTC)
            lTC.Caption = aFechaTC
        End If
                
        If cDisponibilidad.ListCount = 0 Then
            dis_CargoDisponibilidades cDisponibilidad, cMoneda.ItemData(cMoneda.ListIndex)
        End If
        
        Foco tIOriginal
    End If
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    orCheque.fnc_Start cBase
    dis_StartArray
    
    ObtengoSeteoForm Me, Me.Left, Me.Top, Me.Width, Me.Height
    Me.Height = 5700
    
    sNuevo = False: sModificar = False
    InicializoGrillas
    
    cons = "Select MonCodigo, MonSigno from Moneda Where MonCodigo In (" & paMonedaDolar & ", " & paMonedaPesos & ")"
    CargoCombo cons, cMoneda

    DeshabilitoIngreso
    LimpioFicha
    
    FechaDelServidor
    
    If Trim(Command()) <> "" Then CargoCamposDesdeBD Val(Command())
    If prmIdGasto <> 0 Then Botones True, True, True, False, False, Toolbar1, Me
    
    If bPagoConDC Then
        MsgBox "El pago es en moneda extrenjera y acumula diferencias de cambio." & vbCrLf & _
                    "Para visualizar o modificar los datos utilice el formulario de Pagos en Dólares.", vbInformation, "Pago en M/E con Diferencias de Cambio"
        Botones True, False, False, False, False, Toolbar1, Me
    End If
    
    Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
    Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
    Status.Panels("bd") = "BD: " & PropiedadesConnect(prmKeyConnect, Database:=True) & " "

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    GuardoSeteoForm Me
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
End Sub

Private Sub Label1_Click()
    Foco tFSerie
End Sub

Private Sub Label10_Click()
    Foco tComentario
End Sub

Private Sub Label11_Click()
    Foco tID
End Sub

Private Sub Label2_Click()
    Foco tFPaga
End Sub

Private Sub Label3_Click()
    Foco tProveedor
End Sub

Private Sub Label4_Click()
    Foco tFecha
End Sub

Private Sub Label5_Click()
    Foco tSerie
End Sub

Private Sub Label8_Click()
    Foco cMoneda
End Sub

Private Sub Label9_Click()
    Foco tTCDolar
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuEliminar_Click()
    AccionEliminar
End Sub

Private Sub MnuExit_Click()
    Unload Me
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

Sub AccionNuevo(Optional DesdeNuevo As Boolean = False)

    sNuevo = True
    Botones False, False, False, True, True, Toolbar1, Me
    HabilitoIngreso

    If DesdeNuevo Then
        InicializoMData
        tNumero.Text = "": tIOriginal.Text = ""
        tFSerie.Text = "": tFNumero.Text = ""
        tComentario.Text = ""
        
        cDisponibilidad.Text = "": cDisponibilidad.Tag = 0
        orCheque.fnc_BlankControls
        orCheque.Tag = 0

        vsLista.Rows = 1
        Foco tSerie
    Else
        LimpioFicha
        tFecha.Text = Format(Now, "dd/mm/yyyy")
        Foco tFecha
    End If
        
    prmIdGasto = 0
    
End Sub

Private Sub AccionModificar()

    On Error Resume Next
    Screen.MousePointer = 11
    'If Not ValidoDatosMovimientos(prmIdGasto) Then Exit Sub
        
    LimpioFicha
    CargoCamposDesdeBD prmIdGasto
    If prmIdGasto = 0 Then Screen.MousePointer = 0: Exit Sub
    
    sModificar = True
    Botones False, False, False, True, True, Toolbar1, Me
    HabilitoIngreso
    
    If tFecha.Enabled Then Foco tFecha Else Foco tSerie
    Screen.MousePointer = 0
    
End Sub
Private Sub AccionGrabar()

Dim aError As String: aError = ""
Dim bNuevoIngreso As Boolean: bNuevoIngreso = False

    Screen.MousePointer = 11
    If Not ValidoCampos Then Screen.MousePointer = 0: Exit Sub
    If Not ValidoDocumento Then Screen.MousePointer = 0: Exit Sub
    
    Screen.MousePointer = 0
    If MsgBox("Confirma almacenar la información ingresada", vbQuestion + vbYesNo, "Grabar Recibo") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    
    FechaDelServidor
    
    Dim mCompra As Long: mCompra = 0
    If sModificar Then mCompra = prmIdGasto
    If sNuevo Then bNuevoIngreso = True
    
    On Error GoTo errorBT
    cBase.BeginTrans    'COMIENZO TRANSACCION----------------------------------------------!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
        
    'Tabla COMPRA   -----------------------------------------------------------------------------------------------------------
    cons = "Select * from Compra Where ComCodigo = " & mCompra
    Set rsCom = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If sModificar Then
        If gFModificacion <> rsCom!ComFModificacion Then
            aError = "El gasto ha sido modificado por otro usuario. Vuelva a cargar los datos"
            GoTo errorET: Exit Sub
        End If
        rsCom.Edit
    Else
        rsCom.AddNew
    End If
    CargoCamposBDComprobante
    rsCom.Update: rsCom.Close
    '-------------------------------------------------------------------------------------------------------------------------------
    
    If sNuevo Then 'Saco el ID del Recibo
        cons = "Select Max(ComCodigo) from Compra" & _
                    " Where ComTipoDocumento = " & TipoDocumento.CompraReciboDePago & _
                    " And ComProveedor = " & Val(tProveedor.Tag) & _
                    " And ComFecha = '" & Format(tFecha.Text, "mm/dd/yyyy") & "'"
        Set rsCom = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
        mCompra = rsCom(0)
        rsCom.Close
    End If
        
    If tIOriginal.Enabled Then
        'Elimino tabla: Pagos y actualizo los saldos de las compras
        If sModificar Then EliminoPagos mCompra
        CargoCamposBDCompraPago mCompra     'Cargo tabla: CompraPago
    
        GraboBDGastosSubRubro mCompra, sModificar
    End If
    
    If cDisponibilidad.Enabled Then GraboElPago mCompra
    
    If sNuevo Then ActualizoDivisaPaga
        
    cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------
        
    
    sNuevo = False: sModificar = False
    prmIdGasto = mCompra
    gFModificacion = gFechaServidor
    
    DeshabilitoIngreso
    Botones True, True, True, False, False, Toolbar1, Me
    Foco tFecha
    
    If mData.Disponibilidad = 0 And mData.oIDMovimiento = 0 Then
        If MsgBox("Desea ingresar Con Que Paga el Gasto ?.", vbQuestion + vbYesNo, "Ingresa el Pago ?") = vbYes Then
            EjecutarApp prmPathApp & "Con Que Paga.exe", CStr(prmIdGasto)
        End If
    End If
    
    If bNuevoIngreso Then AccionNuevo True
    Screen.MousePointer = 0
    Exit Sub
    
errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    If Trim(aError) = "" Then aError = "No se ha podido inicializar la transacción. Reintente la operación."
    Screen.MousePointer = 0
    clsGeneral.OcurrioError aError, Err.Description
End Sub

Private Sub GraboBDGastosSubRubro(idCompra As Long, Optional Elimino As Boolean = False)

Dim AlSubrubro As Long

    If Elimino Then
        cons = "Delete GastoSubrubro Where GSrIDCompra = " & idCompra
        cBase.Execute cons
    End If
    
    AlSubrubro = paSubrubroAcreedoresVarios
    'Busco si el proveedor tiene un subrubro para asignarlo, sino queda al SR AcreedoresVarios
    cons = " Select * from EmpresaDato" & _
               " Where EDaCodigo = " & Val(tProveedor.Tag) & _
               " And EDaTipoEmpresa = " & TipoEmpresa.Cliente
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then If Not IsNull(rsAux!EDaSRubroContable) Then AlSubrubro = rsAux!EDaSRubroContable
    rsAux.Close
    
    cons = "Select * from GastoSubrubro Where GSrIDCompra = " & idCompra
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    rsAux.AddNew
    rsAux!GSrIDCompra = idCompra
    rsAux!GSrIDSubrubro = AlSubrubro
    rsAux!GSrImporte = CCur(tIOriginal.Text)
    rsAux.Update: rsAux.Close
    
End Sub

Private Sub EliminoPagos(aRecibo As Long)

    'Actualizo los saldos de las facturas
    cons = "Select * from CompraPago Where CPaDocQSalda = " & aRecibo
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not rsAux.EOF
        cons = "Update Compra Set ComSaldo = ComSaldo + " & rsAux!CPaAmortizacion & ", " _
                                            & " ComFModificacion = '" & Format(gFechaServidor, sqlFormatoFH) & "'" _
                & " Where ComCodigo = " & rsAux!CPaDocASaldar
                                            
        cBase.Execute cons
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    'Elimino los pagos ingresados
    cons = "Delete CompraPago Where CPaDocQSalda = " & aRecibo
    cBase.Execute cons
    
End Sub

Private Sub ActualizoDivisaPaga()

    Dim rs1 As rdoResultset
    With vsLista
    
    For I = 1 To .Rows - 1
        If .Cell(flexcpValue, I, 6) - .Cell(flexcpValue, I, 5) = 0 Then
            cons = "Select * from GastoImportacion " _
                   & " Where GImIDCompra = " & .Cell(flexcpValue, I, 4) _
                   & " And GImIDSubrubro = " & paSubrubroDivisa _
                   & " And GImNivelFolder = " & 2
            Set rs1 = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            If Not rs1.EOF Then
                cons = " Update Embarque Set EmbDivisaPaga = 1 Where EmbID = " & rs1!GImFolder
                cBase.Execute cons
            End If
            rs1.Close
        End If
    Next
    End With

End Sub

Private Sub AccionEliminar()
Dim aError As String
    
    If Not ValidoDatosMovimientos(prmIdGasto) Then Exit Sub
    
    If MsgBox("Confirma eliminar el recibo de pago seleccionado", vbQuestion + vbYesNo, "Eliminar") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    
    On Error GoTo errorBT
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
    On Error GoTo errorET
    
    cons = "Select * from Compra Where ComCodigo = " & prmIdGasto
    Set rsCom = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If gFModificacion <> rsCom!ComFModificacion Then
        aError = "El comprobante ha sido modificado recientemente por otro usuario." & vbCrLf & _
                     "Vuelva a cargar los datos"
        GoTo errorET: Exit Sub
    End If
    
    EliminoPagos prmIdGasto 'Elimino tabla: Registro de pagos
    
    cons = "Delete GastoSubrubro Where GSrIDCompra = " & prmIdGasto     'Elimino relacion GastosSubrubro------
    cBase.Execute cons

    rsCom.Delete: rsCom.Close
    
    cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------
    
    LimpioFicha
    DeshabilitoIngreso
    Botones True, False, False, False, False, Toolbar1, Me
    prmIdGasto = 0
    Screen.MousePointer = 0

    Exit Sub
    
errorBT:
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
    Screen.MousePointer = 0: Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionCancelar()
On Error Resume Next

    LimpioFicha
    If sModificar Then
        Botones True, True, True, False, False, Toolbar1, Me
        CargoCamposDesdeBD prmIdGasto
    Else
        Botones True, False, False, False, False, Toolbar1, Me
    End If
    
    DeshabilitoIngreso
    sNuevo = False: sModificar = False
    Foco tFecha

End Sub

Private Sub CargoCamposBDComprobante()

    rsCom!ComTipoDocumento = TipoDocumento.CompraReciboDePago
    rsCom!ComFecha = Format(tFecha.Text, "mm/dd/yyyy")
    rsCom!ComProveedor = Val(tProveedor.Tag)
    rsCom!ComMoneda = cMoneda.ItemData(cMoneda.ListIndex)
    
    If Trim(tSerie.Text) <> "" Then rsCom!ComSerie = Trim(tSerie.Text) Else rsCom!ComSerie = Null
    If Trim(tNumero.Text) <> "" Then rsCom!ComNumero = tNumero.Text Else rsCom!ComNumero = Null
    
    rsCom!ComImporte = CCur(tIOriginal.Text)
    rsCom!ComTC = CCur(tTCDolar.Text)
    
    If Trim(tComentario.Text) <> "" Then rsCom!ComComentario = Trim(tComentario.Text) Else rsCom!ComComentario = Null
    
    rsCom!ComFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    rsCom!ComUsuario = paCodigoDeUsuario
    
    rsCom!ComSaldo = 0
    
End Sub

Private Sub CargoCamposBDCompraPago(aRecibo As Long)

    With vsLista
    
    For I = 1 To .Rows - 1
        'Achico el saldo con lo que se paga
        cons = " Update Compra Set " _
                & " ComSaldo = ComSaldo - " & .Cell(flexcpValue, I, 5) & ", " _
                & " ComFModificacion = '" & Format(gFechaServidor, sqlFormatoFH) & "'" _
                & " Where ComCodigo = " & .Cell(flexcpValue, I, 4)
        cBase.Execute cons
        
        'Grabo relacion Tabla: CompraPago
        cons = "Insert into CompraPago (CPaDocASaldar, CPaDocQSalda, CPaAmortizacion) " _
                & "Values (" & .Cell(flexcpValue, I, 4) & ", " & aRecibo & ", " & .Cell(flexcpValue, I, 5) & ")"
        cBase.Execute cons
    Next
    End With
    
End Sub

Private Sub CargoCamposDesdeBD(idCompra As Long)

Dim aValor As Long

    Screen.MousePointer = 11
    On Error GoTo errCargar
    
    InicializoMData
    
    bPagoConDC = False
    'Cargo los datos desde la tabla COMPRA-----------------------------------------------------------------------------------------
    cons = "Select * from Compra " & _
                " Where ComCodigo = " & idCompra & _
                " And ComTipoDocumento = " & TipoDocumento.CompraReciboDePago
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If rsAux.EOF Then
        rsAux.Close
        MsgBox "No existe registro de un recibo de pago para el id ingresado.", vbExclamation, "No Hay Datos"
        Screen.MousePointer = 0: prmIdGasto = 0
        Botones True, False, False, False, False, Toolbar1, Me
        Exit Sub
    End If
        
    prmIdGasto = rsAux!ComCodigo
    gFModificacion = rsAux!ComFModificacion
    
    tID.Text = Format(rsAux!ComCodigo, "#,##0")
    tFecha.Text = Format(rsAux!ComFecha, "dd/mm/yyyy")
        
    Dim rs1 As rdoResultset
    cons = "Select * from ProveedorCliente Where PClCodigo = " & rsAux!ComProveedor
    Set rs1 = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rs1.EOF Then
        tProveedor.Text = Trim(rs1!PClFantasia)
        tProveedor.Tag = rsAux!ComProveedor
    End If
    rs1.Close
    
    If Not IsNull(rsAux!ComSerie) Then tSerie.Text = Trim(rsAux!ComSerie)
    If Not IsNull(rsAux!ComNumero) Then tNumero.Text = rsAux!ComNumero
    
    BuscoCodigoEnCombo cMoneda, rsAux!ComMoneda
    dis_CargoDisponibilidades cDisponibilidad, rsAux!ComMoneda
    
    tIOriginal.Text = Format(rsAux!ComImporte, "#,##0.00")
    If Not IsNull(rsAux!ComTC) Then If rsAux!ComTC <> 1 Then tTCDolar.Text = Format(rsAux!ComTC, "0.000")
    
    If Not IsNull(rsAux!ComComentario) Then tComentario.Text = Trim(rsAux!ComComentario)
    
    mData.oFechaCompra = rsAux!ComFecha
    If Not IsNull(rsAux!ComUsuario) Then mData.oUsuario = rsAux!ComUsuario
    
    rsAux.Close
    
    'Cargo los datos las facturas pagas-----------------------------------------------------------------------------------------
    Dim aBruto As Currency
    cons = "Select * from CompraPago, Compra" _
           & " Where CPaDocQSalda = " & idCompra _
           & " And CPaDocASaldar = ComCodigo"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        With vsLista
            .AddItem ""
            If Not IsNull(rsAux!ComSerie) Then .Cell(flexcpText, .Rows - 1, 0) = Trim(rsAux!ComSerie)
            .Cell(flexcpText, .Rows - 1, 1) = rsAux!ComNumero
            
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!ComFecha, "dd/mm/yyyy")
            
            aBruto = rsAux!ComImporte
            If Not IsNull(rsAux!ComIva) Then aBruto = aBruto + rsAux!ComIva
            If Not IsNull(rsAux!ComCofis) Then aBruto = aBruto + rsAux!ComCofis
            .Cell(flexcpText, .Rows - 1, 3) = Format(aBruto, "#,##0.00")
                        
            .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!ComCodigo, "#,##0")
            .Cell(flexcpText, .Rows - 1, 5) = Format(rsAux!CPaAmortizacion, "#,##0.00")
            
            If Not IsNull(rsAux!ComDCDe) Then bPagoConDC = True
        End With
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    'Busco los movimientos de Disponibilidades para Ver con que se Pagó --------------------------------------------------------
    Dim mIDDPago As Long, mIDCheque As Long
    mIDDPago = -1: mIDCheque = 0
    cons = "Select * from MovimientoDisponibilidad, MovimientoDisponibilidadRenglon" & _
               " Where MDiIDCompra = " & idCompra & _
               " And MDiID = MDRIDMovimiento "
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not rsAux.EOF Then
        mIDDPago = rsAux!MDRIdDisponibilidad
        mData.oIDMovimiento = rsAux!MDiID
        mData.cndFCierreDisponibilidad = CDate("1/1/1900")
        If Not IsNull(rsAux!MDRIdCheque) Then mIDCheque = rsAux!MDRIdCheque
        
        Dim aDate As Date
        Do While Not rsAux.EOF
            aDate = dis_FechaCierre(rsAux!MDRIdDisponibilidad, mData.oFechaCompra)
            If aDate > mData.cndFCierreDisponibilidad Then mData.cndFCierreDisponibilidad = aDate
            rsAux.MoveNext
            If Not rsAux.EOF Then mIDDPago = 0
        Loop
    End If
    rsAux.Close
    
    orCheque.Tag = 0
    If mIDDPago >= 0 Then
        BuscoCodigoEnCombo cDisponibilidad, mIDDPago
        If cDisponibilidad.ListIndex = -1 Then 'Se pago con otra Disponibilidad,  <> a la MonedadelGasto o Varias Disp.
            BuscoCodigoEnCombo cDisponibilidad, 0
        Else
            If mIDDPago > 0 And mIDCheque <> 0 Then
                With orCheque
                    .prp_IdCheque = mIDCheque
                    .prp_IdDisponibilidad = mIDDPago
                    .prp_IdGasto = idCompra
                    .prp_ValorAAsignar = CCur(tIOriginal.Text)
                    .fnc_Show
                    .Tag = mIDCheque
                End With
            End If
        End If
        
    Else
        cDisponibilidad.Text = ""
        orCheque.fnc_BlankControls
    End If
    
    If cDisponibilidad.ListIndex <> -1 Then
        If cDisponibilidad.ItemData(cDisponibilidad.ListIndex) = 0 Then mData.cndPagoConOtras = True
    End If
    cDisponibilidad.Tag = 0
    '-------------------------------------------------------------------------------------------------------------------------------------------
 
    Screen.MousePointer = 0
    Exit Sub
errCargar:
    clsGeneral.OcurrioError "Error al cargar los datos del comprobante.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub orCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tComentario
End Sub

Private Sub tComentario_GotFocus()
    tComentario.SelStart = 0: tComentario.SelLength = Len(tComentario.Text)
End Sub

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub


Private Sub tFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And Not sNuevo And Not sModificar Then AccionListaDeAyuda
    If KeyCode = vbKeyDown Then tFecha.Text = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub AccionListaDeAyuda()

    On Error GoTo errAyuda
    
    If Not IsDate(tFecha.Text) And Val(tProveedor.Tag) = 0 Then Exit Sub
    Screen.MousePointer = 11
    
    Dim aLista As New clsListadeAyuda
    Dim aIDSel As Long: aIDSel = 0
    
    cons = " Select ID_Compra = ComCodigo, Fecha = ComFecha, Proveedor = PClFantasia, Comprobante = ComSerie + Convert(char(10), ComNumero), Moneda = MonSigno , Importe = ComImporte, Comentarios = ComComentario" _
            & " from Compra, ProveedorCliente, Moneda" _
            & " Where ComProveedor = PClCodigo" _
            & " And ComMoneda = MonCodigo " _
            & " And ComTipoDocumento = " & TipoDocumento.CompraReciboDePago
            
    If IsDate(tFecha.Text) Then cons = cons & " And ComFecha >= '" & Format(tFecha.Text, sqlFormatoF) & "'"
    If Val(tProveedor.Tag) <> 0 Then cons = cons & " And ComProveedor = " & Val(tProveedor.Tag)
    
    cons = cons & " Order by ComFecha ASC"
    
    aIDSel = aLista.ActivarAyuda(cBase, cons, 9000, , "Lista de Recibos")
    If aIDSel <> 0 Then aIDSel = aLista.RetornoDatoSeleccionado(0)
    Me.Refresh
    Set aLista = Nothing
    
    
    If aIDSel <> 0 Then
        LimpioFicha
        CargoCamposDesdeBD aIDSel
    End If
    
    If prmIdGasto <> 0 Then Botones True, True, True, False, False, Toolbar1, Me
    
    If bPagoConDC Then
        MsgBox "El pago es en moneda extrenjera y acumula diferencias de cambio." & vbCrLf & _
                    "Para visualizar o modificar los datos utilice el formulario de Pagos en Dólares.", vbInformation, "Pago en M/E con Diferencias de Cambio"
        Botones True, False, False, False, False, Toolbar1, Me
    End If
    Screen.MousePointer = 0
    Exit Sub
        
errAyuda:
    clsGeneral.OcurrioError "Error al activar la lista de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tFNumero_Change()
    tFNumero.Tag = 0
    lFImporte.Caption = "": lFFecha.Caption = "": lFSaldo.Caption = ""
End Sub

Private Sub tFNumero_GotFocus()
    Status.Panels(4).Text = "F1- Lista de facturas pendientes."
End Sub

Private Sub tFNumero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then Call tFSerie_KeyDown(vbKeyF1, False)
End Sub

Private Sub tFNumero_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Val(tProveedor.Tag) = 0 Then MsgBox "Ingrese el proveedor del recibo de pago.", vbExclamation, "Faltan Datos": Foco tProveedor: Exit Sub
        If cMoneda.ListIndex = -1 Then MsgBox "Seleccionela moneda del recibo de pago.", vbExclamation, "Faltan Datos": Foco cMoneda: Exit Sub
        If Not IsNumeric(tIOriginal.Text) Then MsgBox "Ingrese el importe del recibo de pago.", vbExclamation, "Faltan Datos": Foco tProveedor: Exit Sub
        
        If Val(tFNumero.Tag) <> 0 Then
            tFPaga.Text = Format(CCur(tIOriginal.Text) - HayPago, "#,##0.00")
            Foco tFPaga: Exit Sub
        End If
        
        If Trim(tSerie.Text) = "" And Trim(tNumero.Text) = "" And vsLista.Rows > 1 Then Foco cDisponibilidad: Exit Sub
        
        If Not IsNumeric(tFNumero.Text) Then MsgBox "Ingrese el número de la factura que se paga.", vbExclamation, "Faltan Datos": Foco tFNumero: Exit Sub
        
        On Error GoTo errBusco
        Screen.MousePointer = 11
        'Busco la factura
        
        cons = "Select ID_Compra = ComCodigo, Fecha = ComFecha, Importe = ComImporte, ComIva 'I.V.A', Saldo = ComSaldo, Comentarios = ComComentario from Compra " _
                & " Where ComProveedor = " & Val(tProveedor.Tag) _
                & " And ComTipoDocumento = " & TipoDocumento.CompraCredito _
                & " And ComNumero = " & Trim(tFNumero.Text) _
                & " And ComMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
        If Trim(tFSerie.Text) <> "" Then cons = cons & " And ComSerie = '" & Trim(tFSerie.Text) & "'"

        Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
        Dim aFSeleccionada As Long

        If Not rsAux.EOF Then
            aFSeleccionada = rsAux(0)
            'Valido si hay mas de una factura-------------------------------------------------------------
            rsAux.MoveNext
            If Not rsAux.EOF Then
                rsAux.Close
                aFSeleccionada = ListaDeFacturas(cons)
                If aFSeleccionada = 0 Then Screen.MousePointer = 0: Exit Sub
            Else
                rsAux.Close
            End If
            
            CargoDatosFactura aFSeleccionada
            
        Else
            MsgBox "No existe una factura para el proveedor, moneda y número de documento ingresado.", vbInformation, "No Hay Datos"
            rsAux.Close
        End If
        
        Screen.MousePointer = 0
    End If
    
    Exit Sub
errBusco:
    clsGeneral.OcurrioError "Error al buscar la factura ingresada.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoDatosFactura(Codigo As Long)

    cons = "Select * from Compra Where ComCodigo = " & Codigo & " And ComTipoDocumento = " & TipoDocumento.CompraCredito
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If rsAux.EOF Then
        rsAux.Close
        MsgBox "El comprobante seleccionado no es del tipo Crédito. Verifique.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0: Exit Sub
    Else
        If cMoneda.ItemData(cMoneda.ListIndex) <> rsAux!ComMoneda Then
            rsAux.Close
            MsgBox "La moneda del comprobante seleccionado es distinta a la del recibo de pago. Verifique.", vbExclamation, "ATENCIÓN"
            Screen.MousePointer = 0: Exit Sub
        End If
    End If
    '----------------------------------------------------------------------------------------------------

    If Not IsNull(rsAux!ComSerie) Then tFSerie.Text = Trim(rsAux!ComSerie)
    If Not IsNull(rsAux!ComNumero) Then tFNumero.Text = rsAux!ComNumero
    tFNumero.Tag = rsAux!ComCodigo
    
    Dim aBruto As Currency
    aBruto = rsAux!ComImporte
    If Not IsNull(rsAux!ComIva) Then aBruto = aBruto + rsAux!ComIva
    If Not IsNull(rsAux!ComCofis) Then aBruto = aBruto + rsAux!ComCofis
    lFImporte.Caption = Format(aBruto, "#,##0.00")
    
    lFFecha.Caption = Format(rsAux!ComFecha, "dd/mm/yyyy")
        
    If Not IsNull(rsAux!ComSaldo) Then
        lFSaldo.Caption = Format(rsAux!ComSaldo, "#,##0.00")
        If rsAux!ComSaldo = 0 Then
            MsgBox "El saldo de la factura seleccionada es cero." & vbCrLf & _
                        "Verifique los pagos de ésta factura antes de continuar.", vbExclamation, "Saldo Cero"
        Else
            
            'Verifico si hay ingreso de vencimientos para cargar el valor de una cta.
            Dim rs1 As rdoResultset
            Dim hayVencimientos As Boolean: hayVencimientos = False
            Dim hayPagos As Integer: hayPagos = 0
            Dim aImporteAP As Currency: aImporteAP = 0
            
            cons = "Select * from CompraVencimiento Where CVeIDCompra = " & Codigo
            Set rs1 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
            If Not rs1.EOF Then hayVencimientos = True
            rs1.Close
            
            If hayVencimientos Then
                cons = "Select Count(*) from CompraPago Where CPaDocASaldar = " & Codigo
                Set rs1 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
                If Not rs1.EOF Then If Not IsNull(rs1(0)) Then hayPagos = rs1(0)
                rs1.Close
                
                cons = "Select * from CompraVencimiento Where CVeIDCompra = " & Codigo
                Set rs1 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
                Dim aRow As Integer: aRow = 1
                Do While Not rs1.EOF
                    If hayPagos < aRow Then
                        aImporteAP = rs1!CVeImporte
                        Exit Do
                    Else
                        aRow = aRow + 1
                    End If
                    rs1.MoveNext
                Loop
                rs1.Close
                
                If aImporteAP > 0 Then tFPaga.Text = Format(aImporteAP, "#,##0.00")
            End If
            
            If aImporteAP = 0 Then
                If CCur(tIOriginal.Text) - HayPago > rsAux!ComSaldo Then
                    tFPaga.Text = Format(rsAux!ComSaldo, "#,##0.00")
                Else
                    tFPaga.Text = Format(CCur(tIOriginal.Text) - HayPago, "#,##0.00")
                End If
            End If
            
        End If
    End If
    rsAux.Close
    Foco tFPaga
            
End Sub

Private Sub tFNumero_LostFocus()
    Status.Panels(4).Text = ""
End Sub

Private Sub tFPaga_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And IsNumeric(tFPaga.Text) Then bAgregar.SetFocus
End Sub

Private Sub tFPaga_LostFocus()
    On Error Resume Next
    If IsNumeric(tFPaga.Text) Then tFPaga.Text = Format(tFPaga.Text, "#,##0.00")
End Sub

Private Sub tFSerie_Change()
    tFNumero.Tag = 0
    lFImporte.Caption = "": lFFecha.Caption = "": lFSaldo.Caption = ""
End Sub

Private Sub tFSerie_GotFocus()
    tFSerie.SelStart = 0: tFSerie.SelLength = Len(tFSerie.Text)
    Status.Panels(4).Text = "F1- Lista de facturas pendientes."
End Sub

Private Sub tFSerie_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then
        If Val(tProveedor.Tag) = 0 Then MsgBox "Ingrese el proveedor del recibo de pago.", vbExclamation, "ATENCIÓN": Foco tProveedor: Exit Sub
        If cMoneda.ListIndex = -1 Then MsgBox "Seleccionela moneda del recibo de pago.", vbExclamation, "ATENCIÓN": Foco cMoneda: Exit Sub
        If Not IsNumeric(tIOriginal.Text) Then MsgBox "Ingrese el importe del recibo de pago.", vbExclamation, "ATENCIÓN": Foco tProveedor: Exit Sub
        
        On Error GoTo errBusco
        Screen.MousePointer = 11
        
        cons = "Select ID_Compra = ComCodigo, Fecha = ComFecha, Serie = ComSerie, ComNumero 'Número', Importe = ComImporte, ComIva 'I.V.A', Saldo = ComSaldo, Comentarios = ComComentario from Compra " _
                & " Where ComProveedor = " & CLng(tProveedor.Tag) _
                & " And ComTipoDocumento = " & TipoDocumento.CompraCredito _
                & " And ComMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                & " And ComSaldo > 0 "
        
        Dim aFSeleccionada As Long
        aFSeleccionada = ListaDeFacturas(cons)
        If aFSeleccionada = 0 Then Screen.MousePointer = 0: Exit Sub
        CargoDatosFactura aFSeleccionada
        
        Screen.MousePointer = 0
    End If
    
    Exit Sub
errBusco:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al buscar las facturas del proveedor.", Err.Description
End Sub

Private Sub tFSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then Foco tFNumero
End Sub

Private Sub tFSerie_LostFocus()
    Status.Panels(4).Text = ""
End Sub

Private Sub tID_Change()
    If tID.Enabled Then Botones True, False, False, False, False, Toolbar1, Me
End Sub

Private Sub tID_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Trim(tID.Text) = "" Then Foco tFecha: Exit Sub
        If Not IsNumeric(tID.Text) Then MsgBox "El id ingresado no es correcto. Verifique.", vbExclamation, "ATENCIÓN": Exit Sub
        prmIdGasto = CLng(tID.Text)
        LimpioFicha
        CargoCamposDesdeBD prmIdGasto
        If prmIdGasto <> 0 Then Botones True, True, True, False, False, Toolbar1, Me
        
        If bPagoConDC Then
            MsgBox "El pago es en moneda extrenjera y acumula diferencias de cambio." & Chr(vbKeyReturn) & "Para visualizar o modificar los datos utilice el formulario de Pagos en Dólares.", vbInformation, "Pago en M/E con Diferencias de Cambio"
            Botones True, False, False, False, False, Toolbar1, Me
        End If
            
    End If
    
End Sub

Private Sub tNumero_GotFocus()
    tNumero.SelStart = 0: tNumero.SelLength = Len(tNumero.Text)
End Sub

Private Sub tNumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cMoneda.Enabled Then Foco cMoneda: Exit Sub
        If tIOriginal.Enabled Then Foco tIOriginal Else Foco tTCDolar
    End If
End Sub

Private Sub tFecha_GotFocus()
    tFecha.SelStart = 0
    tFecha.SelLength = Len(tFecha.Text)
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then If tProveedor.Enabled Then Foco tProveedor Else Foco tSerie
End Sub

Private Sub tFecha_LostFocus()
    If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy") Else tFecha.Text = ""
End Sub

Private Sub tIOriginal_GotFocus()
    tIOriginal.SelStart = 0
    tIOriginal.SelLength = Len(tIOriginal.Text)
End Sub

Private Sub tIOriginal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tFSerie
End Sub

Private Sub tIOriginal_LostFocus()

    If Not IsNumeric(tIOriginal.Text) Then tIOriginal.Text = ""
    tIOriginal.Text = Format(tIOriginal.Text, "##,##0.00")
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        
        Case "pagos": EjecutarApp prmPathApp & "Con Que Paga.exe", CStr(prmIdGasto)
        Case "dolar": EjecutarApp prmPathApp & "Tasa de Cambio.exe"
        
        Case "salir": Unload Me
    End Select

End Sub

Private Sub DeshabilitoIngreso()

    tID.BackColor = Blanco: tID.Enabled = True
    tFecha.BackColor = Blanco: tFecha.Enabled = True
    
    tProveedor.Enabled = True: tProveedor.BackColor = Blanco
    tSerie.Enabled = False: tSerie.BackColor = Inactivo
    tNumero.Enabled = False: tNumero.BackColor = Inactivo
    
    cMoneda.Enabled = False: cMoneda.BackColor = Inactivo
    tIOriginal.Enabled = False: tIOriginal.BackColor = Inactivo
    tTCDolar.Enabled = False: tTCDolar.BackColor = Inactivo
    
    tFSerie.Enabled = False: tFSerie.BackColor = Inactivo
    tFNumero.Enabled = False: tFNumero.BackColor = Inactivo
    tFPaga.Enabled = False: tFPaga.BackColor = Inactivo
    
    tComentario.Enabled = False: tComentario.BackColor = Inactivo
    vsLista.BackColor = Inactivo
    
    bAgregar.Enabled = False
    
    cDisponibilidad.Enabled = False: cDisponibilidad.BackColor = Colores.Inactivo
    orCheque.Enabled = False: orCheque.BackColor = Colores.Inactivo
    
End Sub

Private Sub HabilitoIngreso()
    
    ' NO SE Puden Modificar: Proveedor, Moneda, Importe, Facturas asociadas.
    
    mData.flgSucesoXMod = False
    tID.BackColor = Inactivo: tID.Enabled = False
    
    If sModificar Then
        tProveedor.BackColor = Colores.Inactivo: tProveedor.Enabled = False
        cMoneda.BackColor = Colores.Inactivo: cMoneda.Enabled = False
    
        tFecha.Enabled = True
        If mData.oFechaCompra <= prmFCierreIVA Then tFecha.Enabled = False: tFecha.BackColor = Colores.Inactivo
        If mData.oFechaCompra <= mData.cndFCierreDisponibilidad Then tFecha.Enabled = False: tFecha.BackColor = Colores.Inactivo
        
        Dim bADM As Boolean
    
        If mData.oFechaCompra < prmFCierreIVA Then      'ANTES DEL CIERRE DEL IVA
            mData.flgSucesoXMod = True  'Si se hizo el cierre del iva va suceso siempre
            bADM = miConexion.AccesoAlMenu(prmKeyAppADM)
            
            If bADM Then
                tSerie.Enabled = True: tSerie.BackColor = vbWindowBackground
                tNumero.Enabled = True: tNumero.BackColor = vbWindowBackground
                tTCDolar.Enabled = True: tTCDolar.BackColor = Colores.Obligatorio
                tComentario.Enabled = True: tComentario.BackColor = vbWindowBackground
            End If
            
        ElseIf mData.oFechaCompra > mData.cndFCierreDisponibilidad Then     'MAYOR AL CIERRE DE LA DISP.
                    
            tSerie.Enabled = True: tSerie.BackColor = vbWindowBackground
            tNumero.Enabled = True: tNumero.BackColor = vbWindowBackground
            'tIOriginal.Enabled = True: tIOriginal.BackColor = Colores.Obligatorio
            tTCDolar.Enabled = True: tTCDolar.BackColor = Colores.Obligatorio
        
            tComentario.Enabled = True: tComentario.BackColor = vbWindowBackground
            cDisponibilidad.Enabled = True: cDisponibilidad.BackColor = vbWindowBackground
        
        Else        'ANTES DEL CIERRE DE LA DISPONIBILIDAD
            tSerie.Enabled = True: tSerie.BackColor = vbWindowBackground
            tNumero.Enabled = True: tNumero.BackColor = vbWindowBackground
            tTCDolar.Enabled = True: tTCDolar.BackColor = Colores.Obligatorio
            tComentario.Enabled = True: tComentario.BackColor = vbWindowBackground
        End If
        
        'Si se pagó con una Disp con moneda dif a la del gasto o con varias ...
        '1) No dejo modificar Disponibilidad, Comprobante y Monedas
        If mData.cndPagoConOtras Or (mData.oFechaCompra <= mData.cndFCierreDisponibilidad) Then
            cDisponibilidad.Enabled = False: cDisponibilidad.BackColor = Colores.Inactivo
            cMoneda.Enabled = False: cMoneda.BackColor = Colores.Inactivo
        End If
    
    Else
    
        tFecha.BackColor = Obligatorio
        tProveedor.BackColor = Obligatorio
        tSerie.Enabled = True: tSerie.BackColor = Blanco
        tNumero.Enabled = True: tNumero.BackColor = Obligatorio
    
        cMoneda.Enabled = True: cMoneda.BackColor = Obligatorio
        tIOriginal.Enabled = True: tIOriginal.BackColor = Obligatorio
        tTCDolar.Enabled = True: tTCDolar.BackColor = Obligatorio
    
        tFSerie.Enabled = True: tFSerie.BackColor = Blanco
        tFNumero.Enabled = True: tFNumero.BackColor = Blanco
        tFPaga.Enabled = True: tFPaga.BackColor = Blanco
    
        tComentario.Enabled = True: tComentario.BackColor = Blanco
        vsLista.BackColor = Blanco
    
        bAgregar.Enabled = True
        
        cDisponibilidad.Enabled = True: cDisponibilidad.BackColor = vbWindowBackground
    End If
        
End Sub

Private Sub LimpioFicha()
    
    InicializoMData
    
    tID.Text = ""
    tFecha.Text = ""
    tProveedor.Text = ""
    cMoneda.Text = "": tIOriginal.Text = ""
    tSerie.Text = "": tNumero.Text = ""
    
    tTCDolar.Text = ""
    lTC.Caption = ""
    
    tFSerie.Text = "": tFNumero.Text = "": tFPaga.Text = ""
    lFImporte.Caption = "": lFFecha.Caption = "": lFSaldo.Caption = ""
    
    vsLista.Rows = 1
    tComentario.Text = ""
    
    cDisponibilidad.Text = "": cDisponibilidad.Tag = 0
    orCheque.fnc_BlankControls
    orCheque.Tag = 0
    
    
End Sub

Private Sub tProveedor_Change()
    tProveedor.Tag = 0
End Sub

Private Sub tProveedor_GotFocus()
    tProveedor.SelStart = 0: tProveedor.SelLength = Len(tProveedor.Text)
End Sub

Private Sub tProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And Val(tProveedor.Tag) <> 0 Then AccionListaDeAyuda
End Sub

Private Sub tProveedor_KeyPress(KeyAscii As Integer)
    On Error GoTo errBuscar
    
    If KeyAscii = vbKeyReturn Then
        If Val(tProveedor.Tag) <> 0 Or Trim(tProveedor.Text) = "" Then If tID.Enabled Then Foco tID Else Foco tSerie: Exit Sub
        Screen.MousePointer = 11
        
        Screen.MousePointer = 11
        Dim aQ As Long, aIdProveedor As Long, aTexto As String
        
        aQ = 0
        cons = "Select PClCodigo, PClFantasia as 'Nombre', PClNombre 'Razón Social' " & _
                    " From ProveedorCliente " & _
                    " Where PClNombre like '" & Trim(tProveedor.Text) & "%' Or PClFantasia like '" & Trim(tProveedor.Text) & "%'"
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            aQ = 1: aIdProveedor = rsAux!PClCodigo: aTexto = Trim(rsAux(1))
            rsAux.MoveNext: If Not rsAux.EOF Then aQ = 2
        End If
        rsAux.Close
        
        Select Case aQ
            Case 0:
                    MsgBox "No existe una empresa para el con el nombre ingresado.", vbExclamation, "No existe Empresa"
            
            Case 1:
                    tProveedor.Text = aTexto
                    tProveedor.Tag = aIdProveedor
                    If (sNuevo Or sModificar) Then Foco tSerie
        
            Case 2:
                    Dim aLista As New clsListadeAyuda
                    Dim aIDSel As Long, aTxtSel As String
                    
                    aIDSel = aLista.ActivarAyuda(cBase, cons, 5500, 1, "Proveedores")
                    If aIDSel <> 0 Then
                        aIDSel = aLista.RetornoDatoSeleccionado(0)
                        aTxtSel = aLista.RetornoDatoSeleccionado(1)
                    End If
                    Set aLista = Nothing

                    If aIDSel <> 0 Then
                        tProveedor.Text = aTxtSel
                        tProveedor.Tag = aIDSel
                        
                        If (sNuevo Or sModificar) Then Foco tSerie
                    Else
                        tProveedor.Text = ""
                    End If
        End Select
        Screen.MousePointer = 0
    End If
    Exit Sub

errBuscar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al procesar la lista de ayuda.", Err.Description
End Sub

Private Sub tSerie_GotFocus()
    tSerie.SelStart = 0: tSerie.SelLength = Len(tSerie.Text)
End Sub

Private Sub tSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then Foco tNumero
End Sub

Private Sub tTCDolar_Change()
    lTC.Caption = "manual"
End Sub

Private Sub tTCDolar_GotFocus()
    tTCDolar.SelStart = 0: tTCDolar.SelLength = Len(tTCDolar.Text)
End Sub

Private Sub tTCDolar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tFSerie.Enabled Then Foco tFSerie: Exit Sub
        If cDisponibilidad.Enabled Then cDisponibilidad.SetFocus Else Foco tComentario
    End If
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsLista
        .Rows = 1: .Cols = 1
        .Editable = False
        .FormatString = "<Serie|<Numero|Fecha|>Importe|<Id_Compra|>Paga|>Saldo Actual"
        .ExtendLastCol = True
        .WordWrap = True
        .ColWidth(1) = 1100: .ColWidth(1) = 1000: .ColWidth(2) = 1000: .ColWidth(3) = 1300:: .ColWidth(4) = 1000:: .ColWidth(5) = 1300
        .ColDataType(2) = flexDTCurrency: .ColDataType(5) = flexDTCurrency
        .ExtendLastCol = True
    End With
    
End Sub


Private Sub vsLista_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
        If vsLista.RowSel < 1 Then Exit Sub
        If sModificar Then Exit Sub
        If Not sNuevo And Not sModificar Then Exit Sub
        On Error Resume Next
        vsLista.RemoveItem vsLista.RowSel
    End If
    
End Sub

Private Function ValidoCampos() As Boolean

Dim aTotal As Currency
    
    On Error GoTo errValido
    ValidoCampos = False
    
    If Not IsDate(tFecha.Text) Then
        MsgBox "La fecha ingresada para el registro del comprobante no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tFecha: Exit Function
    End If
    If Val(tProveedor.Tag) = 0 Then
        MsgBox "Debe seleccionar el proveedor del pago.", vbExclamation, "ATENCIÓN"
        Foco tProveedor: Exit Function
    End If
    
    If Not IsNumeric(tNumero.Text) Then
        MsgBox "Debe ingresar la numeración del comprobante.", vbExclamation, "ATENCIÓN"
        Foco tNumero: Exit Function
    End If
    
    If cMoneda.ListIndex = -1 Then
        MsgBox "Debe seleccionar la moneda para el registro del pago.", vbExclamation, "ATENCIÓN"
        Foco cMoneda: Exit Function
    End If
    If Not IsNumeric(tIOriginal.Text) Then
        MsgBox "Debe ingresar el importe total del recibo de pago.", vbExclamation, "ATENCIÓN"
        Foco tIOriginal: Exit Function
    End If
    
    If Not IsNumeric(tTCDolar.Text) Then
        MsgBox "Debe ingresar el valor del dólar para la fecha ingresada (tasa de cambio).", vbExclamation, "ATENCIÓN"
        Foco tNumero: Exit Function
    End If
        
    If vsLista.Rows = 1 Then
        MsgBox "Debe ingresar las facturas a pagar con el recibo.", vbExclamation, "ATENCIÓN"
        Foco tFSerie: Exit Function
    End If
       
    'Valido importe del las facturas contra el importe original
    If HayPago <> CCur(tIOriginal.Text) Then
        MsgBox "El importe del comprobante (" & tIOriginal.Text & ") no coincide con la suma de los gastos (" & Format(HayPago, "#,##0.00") & ").", vbExclamation, "ATENCIÓN"
        Foco tIOriginal: Exit Function
    End If
    
    If cDisponibilidad.ListIndex = -1 Then
        MsgBox "Debe seleccionar con que se paga el recibo.", vbExclamation, "Falta Con Que se Paga"
        Foco cDisponibilidad: Exit Function
    End If
    
    'Controlo la fecha del Gasto -------------------------------------------------------------------------------------------------
    If tFecha.Enabled Then
        If cDisponibilidad.ListIndex > 0 Then
            Dim mDate As Date, mDateG As Date
            Dim bCerrada As Boolean, mIdx As Integer
            bCerrada = False
            mDate = dis_FechaCierre(cDisponibilidad.ItemData(cDisponibilidad.ListIndex), CDate(tFecha.Text))
            mDateG = CDate(tFecha.Text)
            
            mIdx = dis_IdxArray(cDisponibilidad.ItemData(cDisponibilidad.ListIndex))
            If arrDisp(mIdx).Bancaria Then
                If orCheque.fnc_GetValorData("") Then   'Hay Cheque, si no a la orden
                    If Trim(orCheque.fnc_GetValorData("CheVencimiento")) <> "" Then
                        mDateG = CDate(orCheque.fnc_GetValorData("CheVencimiento"))
                    Else
                        mDateG = CDate(orCheque.fnc_GetValorData("CheLibrado"))
                    End If
                End If
            End If
            If mDate >= mDateG Then bCerrada = True
            
            If bCerrada Then
                MsgBox "La disponibilidad " & cDisponibilidad.Text & " está cerrada." & vbCrLf & _
                            "No se pueden realizar movimientos para la fecha del gasto.", vbExclamation, "Disponibilidad Cerrada"
                Exit Function
            End If
        End If
        
        If prmFCierreIVA >= CDate(tFecha.Text) Then
            MsgBox "El pago de impuestos está cerrado al " & Format(prmFCierreIVA, "d/mm/yyyy") & "." & vbCrLf & _
                        "No se pueden ingresar gastos menores a esa fecha.", vbExclamation, "Pago de Impuestos Cerrado"
            Exit Function
        End If
    End If
    '-----------------------------------------------------------------------------------------------------------------------------------
    
    'Valido los datos del Cheque 'Importe del Gasto Contra Importe Disponible
    If orCheque.fnc_GetValorData("") Then
        Dim mIDC As Long, mDisponibile As Currency
        mIDC = orCheque.fnc_GetValorData("CheID")
        mDisponibile = orCheque.fnc_GetValorData("CheImporte")
        
        If mIDC <> 0 Then
            mDisponibile = mDisponibile - dis_ImporteAsignadoCheque(mIDC, prmIdGasto)
        End If
        If CCur(tIOriginal.Text) > mDisponibile Then
            MsgBox "El valor del gasto no debe superar el importe disponible del cheque.", vbInformation, "Importe Gasto > al del Cheque"
            Foco tIOriginal: Exit Function
        End If
    End If
    
    'Cargo Estructura mData con valores para usar al grabar ---------------------------------------------------------------------------
    With mData
        .ImporteCompra = CCur(tIOriginal.Text)
        .ImporteDisponibilidad = CCur(tIOriginal.Text)
        .ImportePesos = 0
        
        .Disponibilidad = -1
        If cDisponibilidad.ListIndex <> -1 Then .Disponibilidad = cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
        
        If .Disponibilidad <> -1 And .Disponibilidad <> 0 Then      '-1= Nada; 0= Split
            Dim mMonedaG As Integer    'Voy a sacar el importe en pesos
            mMonedaG = cMoneda.ItemData(cMoneda.ListIndex)
        
            If mMonedaG = paMonedaPesos Then
                .ImportePesos = .ImporteCompra
            Else
                
                If mMonedaG = paMonedaDolar Then
                    .ImportePesos = .ImporteCompra * CCur(tTCDolar.Text)
                Else
                    .ImportePesos = TasadeCambio(mMonedaG, paMonedaPesos, CDate(tFecha.Text), "")
                    .ImportePesos = .ImporteDisponibilidad * .ImportePesos
                End If
            End If
        End If
        'SACO LOS Valores Absolutos x ingreso de pagos en negativo
        .ImporteCompra = (.ImporteCompra)
        .ImporteDisponibilidad = (.ImporteDisponibilidad)
        .ImportePesos = (.ImportePesos)
        
        .HaceSalidaCaja = True
    End With
    '------------------------------------------------------------------------------------------------------------------------------------------------------
    
    ValidoCampos = True
    Exit Function

errValido:
    clsGeneral.OcurrioError "Error al validar los datos.", Err.Description
    Screen.MousePointer = 0
End Function

Private Function ValidoDocumento() As Boolean
    
    On Error Resume Next
    ValidoDocumento = False
    
    cons = "Select * from Compra Where ComCodigo <> " & prmIdGasto
           
    If Trim(tNumero.Text) <> "" Then cons = cons & " And ComNumero = " & Trim(tNumero.Text)
    
    cons = cons & " And ComProveedor = " & Val(tProveedor.Tag) _
                       & " And ComMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                       & " And ComImporte = " & CCur(tIOriginal.Text) _
                       & " And ComTipoDocumento = " & TipoDocumento.CompraReciboDePago
           
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        Screen.MousePointer = 0
        If MsgBox("Ya existen recibos registrados con el mismo documento y proveedor." & Chr(vbKeyReturn) _
            & "Fecha: " & Format(rsAux!ComFecha, "d-mmm yyyy") & Chr(vbKeyReturn) _
            & "Importe: " & Format(rsAux!ComImporte, "##,##0.00") & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
            & "Desea proseguir con el ingreso del gasto.", vbInformation + vbYesNo + vbDefaultButton2, "ATENCIÓN") = vbNo Then
                rsAux.Close: Exit Function
        End If
    End If
    rsAux.Close
    Screen.MousePointer = 0
    ValidoDocumento = True
                   
End Function

Private Function ValidoDatosMovimientos(aRecibo As Long) As Boolean

    On Error GoTo errValidar
    ValidoDatosMovimientos = True
    
    'Valido los campos de la tabla vencimiento-------------------------------------------------------------------------------------------
    cons = "Select * from MovimientoDisponibilidad Where MDiIdCompra = " & aRecibo
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        ValidoDatosMovimientos = False
        Screen.MousePointer = 0
        MsgBox "Hay movimientos de disponibilidades ingresados para el comprobante." & Chr(vbKeyReturn) & "Para continuar con la acción debe eliminarlos.", vbInformation, "Movimientos de Disponibilidades"
    End If
    rsAux.Close
    Exit Function

errValidar:
    clsGeneral.OcurrioError "Ocurrió un error al validar movimientos de disponibilidades.", Err.Description
End Function


Private Function HayPago() As Currency
    Dim aRetorno As Currency
    
    aRetorno = 0
    With vsLista
        For I = 1 To .Rows - 1: aRetorno = aRetorno + .Cell(flexcpValue, I, 5): Next
    End With
    HayPago = aRetorno
    
End Function

Private Function ListaDeFacturas(Consulta As String) As Long

    On Error GoTo errAyuda
    ListaDeFacturas = 0
    
    Dim aLista As New clsListadeAyuda
    Dim aIDSel  As Long: aIDSel = 0
    
    aLista.ActivoListaAyudaSQL cBase, Consulta
    If Trim(aLista.ItemSeleccionadoSQL) <> "" Then aIDSel = aLista.ItemSeleccionadoSQL
    
    Me.Refresh
    Set aLista = Nothing
    
    ListaDeFacturas = aIDSel
    Screen.MousePointer = 0
    Exit Function
        
errAyuda:
    clsGeneral.OcurrioError "Error al activar la lista de ayuda.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub GraboElPago(mIDCompra As Long)
    
    If mData.oIDMovimiento = 0 And (mData.Disponibilidad = -1 Or mData.Disponibilidad = 0) Then Exit Sub
    
    Dim mIDCheque As Long
    mIDCheque = Val(orCheque.Tag)
    
    '1) Si Antes se hizo movimiento y Ahora no hay o pone otras .... Lo Borro
    If mData.oIDMovimiento <> 0 And (mData.Disponibilidad = -1 Or mData.Disponibilidad = 0) Then
        
        cons = "Delete MovimientoDisponibilidadRenglon Where MDRIdMovimiento = " & mData.oIDMovimiento
        cBase.Execute cons
        
        cons = "Delete MovimientoDisponibilidad Where MDiID = " & mData.oIDMovimiento
        cBase.Execute cons
        
        If mIDCheque <> 0 Then dis_BorroRelacionCheque mIDCheque, mIDCompra
        Exit Sub
    End If
    
Dim RsMov As rdoResultset
Dim mFechaHora As String
Dim mIDMov As Long
    
    mFechaHora = Format(tFecha.Text, "dd/mm/yyyy") & " " & Format(gFechaServidor, "hh:mm:ss")
    mIDMov = mData.oIDMovimiento
    
    'Inserto en la Tabla Movimiento-Disponibilidad--------------------------------------------------------
    cons = "Select * from MovimientoDisponibilidad Where MDiID = " & mIDMov
    Set RsMov = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If mIDMov = 0 Then RsMov.AddNew Else RsMov.Edit
    
    RsMov!MDiFecha = Format(mFechaHora, "mm/dd/yyyy")
    RsMov!MDiHora = Format(mFechaHora, "hh:mm:ss")
    RsMov!MDiTipo = paMDPagoDeCompra
    RsMov!MDiIdCompra = mIDCompra
    RsMov!MDiComentario = Trim(tProveedor.Text)
    RsMov.Update: RsMov.Close
    '------------------------------------------------------------------------------------------------------------
    
    'Saco el Id de movimiento-------------------------------------------------------------------------------
    If mIDMov = 0 Then
        cons = "Select Max(MDiID) from MovimientoDisponibilidad" & _
                  " Where MDiFecha = " & Format(mFechaHora, "'mm/dd/yyyy'") & _
                  " And MDiHora = " & Format(mFechaHora, "'hh:mm:ss'") & _
                  " And MDiTipo = " & paMDPagoDeCompra & _
                  " And MDiIdCompra = " & mIDCompra
        
        Set RsMov = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        mIDMov = RsMov(0)
        RsMov.Close
    End If
    '------------------------------------------------------------------------------------------------------------
    
    'Valido Si hay que ingresar Cheque --------------------------------------------------------------------
    If orCheque.fnc_GetValorData("") Then
        Dim mNewCheque As Long
        mNewCheque = orCheque.fnc_GetValorData("CheID")
        If mNewCheque = 0 Then
            If mIDCheque > 0 Then dis_BorroRelacionCheque mIDCheque, mIDCompra  'Borro rel al cheque anterior
            
            'Agrego el nuevo Cheque
            cons = "Select * from Cheque Where CheID = 0"
            Set RsMov = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            RsMov.AddNew
            
            RsMov!CheIDDisponibilidad = mData.Disponibilidad
            RsMov!CheSerie = Trim(orCheque.fnc_GetValorData("CheSerie"))
            RsMov!CheNumero = Trim(orCheque.fnc_GetValorData("CheNumero"))
            RsMov!CheImporte = orCheque.fnc_GetValorData("CheImporte")
            RsMov!CheLibrado = Format(orCheque.fnc_GetValorData("CheLibrado"), "mm/dd/yyyy")
            If Trim(orCheque.fnc_GetValorData("CheVencimiento")) <> "" Then RsMov!CheVencimiento = Format(orCheque.fnc_GetValorData("CheVencimiento"), "mm/dd/yyyy")
                    
            RsMov.Update: RsMov.Close
                    
            cons = "Select Max(CheID) from Cheque" & _
                    " Where CheIDDisponibilidad = " & mData.Disponibilidad
            Set RsMov = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            mIDCheque = RsMov(0)
            RsMov.Close
        
            cons = "Select * from ChequePago Where CPaIDCheque = " & mIDCheque & " And CPaIDCompra = " & mIDCompra
            Set RsMov = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            RsMov.AddNew
            RsMov!CPaIDCheque = mIDCheque
            RsMov!CPaIDCompra = mIDCompra
            RsMov!CPaImporte = mData.ImporteDisponibilidad
            RsMov.Update: RsMov.Close
        
        Else
            'El nuevo cheque ya existe
            If mNewCheque = mIDCheque Then  'Si es el mismo, solametne actualizo la relacion
                cons = "Select * from ChequePago " & _
                            " Where CPaIDCheque = " & mIDCheque & " And CPaIDCompra = " & mIDCompra
                Set RsMov = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                If RsMov!CPaImporte <> mData.ImporteDisponibilidad Then
                    RsMov.Edit
                    RsMov!CPaImporte = mData.ImporteDisponibilidad
                    RsMov.Update
                End If
                RsMov.Close
                
            Else        'Es otro Cheque 1) Borro rel al viejo, 2) hago rel al nuevo
                If mIDCheque <> 0 Then dis_BorroRelacionCheque mIDCheque, mIDCompra
                
                cons = "Select * from ChequePago " & _
                            " Where CPaIDCheque = " & mNewCheque & " And CPaIDCompra = " & mIDCompra
                Set RsMov = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                RsMov.AddNew
                RsMov!CPaIDCheque = mNewCheque
                RsMov!CPaIDCompra = mIDCompra
                RsMov!CPaImporte = mData.ImporteDisponibilidad
                RsMov.Update: RsMov.Close
                
                mIDCheque = mNewCheque
            End If
        End If
    Else
        If mIDCheque > 0 Then dis_BorroRelacionCheque mIDCheque, mIDCompra  'Borro rel al cheque anterior
        mIDCheque = 0
    End If
    '------------------------------------------------------------------------------------------------------------
    
    'Grabo en Tabla Movimiento-Disponibilidad-Renglon--------------------------------------------------
    cons = "Select * from MovimientoDisponibilidadRenglon Where MDRIdMovimiento = " & mIDMov
    Set RsMov = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If RsMov.EOF Then RsMov.AddNew Else RsMov.Edit
    
    RsMov!MDRIdMovimiento = mIDMov
    RsMov!MDRIdDisponibilidad = mData.Disponibilidad
    RsMov!MDRIdCheque = mIDCheque
    
    RsMov!MDRImporteCompra = mData.ImporteCompra
    RsMov!MDRImportePesos = mData.ImportePesos
    
    If mData.HaceSalidaCaja Then
        RsMov!MDRHaber = mData.ImporteDisponibilidad
    Else
        RsMov!MDRDebe = mData.ImporteDisponibilidad
    End If
    
    RsMov.Update: RsMov.Close

End Sub


Private Function InicializoMData()

    With mData
        .oFechaCompra = CDate("01/01/1900")
        .oIDMovimiento = 0
        .oTotalBruto = 0
        .oPesos = 0
        .oUsuario = 0
        
        .Disponibilidad = 0
        .ImporteCompra = 0
        .ImporteDisponibilidad = 0
        .ImportePesos = 0
        .HaceSalidaCaja = False
        
        .cndPagoConOtras = False
        .cndFCierreDisponibilidad = CDate("01/01/1900")
        
        .flgSucesoXMod = False
    End With
    
End Function

