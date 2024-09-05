VERSION 5.00
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmEntrega 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha de Entrega"
   ClientHeight    =   4920
   ClientLeft      =   3375
   ClientTop       =   2460
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
   Icon            =   "frmEntrega.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7245
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   500
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
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
      EndProperty
   End
   Begin VB.PictureBox picRetiro 
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2355
      ScaleWidth      =   7035
      TabIndex        =   29
      Top             =   2100
      Width           =   7095
      Begin VB.CommandButton bFactura 
         Caption         =   "&Ver Factura"
         Height          =   295
         Left            =   1920
         TabIndex        =   42
         Top             =   60
         Width           =   1275
      End
      Begin VB.CheckBox chLiquidada 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6720
         TabIndex        =   41
         Top             =   1140
         Width           =   255
      End
      Begin VB.TextBox tUsuario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   11
         TabIndex        =   17
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox tRFecha 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         MaxLength       =   15
         TabIndex        =   1
         Top             =   420
         Width           =   1335
      End
      Begin VB.TextBox tRComentario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   15
         Top             =   1620
         Width           =   5880
      End
      Begin VB.TextBox tRCosto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   9
         Top             =   1140
         Width           =   1095
      End
      Begin VB.TextBox tRLiquidar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5805
         MaxLength       =   15
         TabIndex        =   13
         Top             =   1170
         Width           =   855
      End
      Begin AACombo99.AACombo cRAsignado 
         Height          =   315
         Left            =   840
         TabIndex        =   4
         Top             =   780
         Width           =   1875
         _ExtentX        =   3307
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
         Text            =   ""
      End
      Begin AACombo99.AACombo cRMoneda 
         Height          =   315
         Left            =   840
         TabIndex        =   8
         Top             =   1140
         Width           =   675
         _ExtentX        =   1191
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
         Text            =   ""
      End
      Begin AACombo99.AACombo cRFPago 
         Height          =   315
         Left            =   3420
         TabIndex        =   11
         Top             =   1140
         Width           =   1455
         _ExtentX        =   2566
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
         Text            =   ""
      End
      Begin AACombo99.AACombo cRFlete 
         Height          =   315
         Left            =   3420
         TabIndex        =   6
         Top             =   780
         Width           =   1455
         _ExtentX        =   2566
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
         Text            =   ""
      End
      Begin AACombo99.AACombo cRHora 
         Height          =   315
         Left            =   2220
         TabIndex        =   2
         Top             =   420
         Width           =   1455
         _ExtentX        =   2566
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
         Text            =   ""
      End
      Begin VB.Label lRModificado 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/09/00 23:55"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5580
         TabIndex        =   40
         Top             =   1920
         Width           =   1365
      End
      Begin VB.Label lRSinEfecto 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SIN EFECTO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3300
         TabIndex        =   37
         Top             =   60
         Width           =   1305
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Co&mentarios:"
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   1620
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "&Usuario:"
         Height          =   195
         Left            =   60
         TabIndex        =   16
         Top             =   1950
         Width           =   735
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Modificado:"
         Height          =   195
         Left            =   4620
         TabIndex        =   36
         Top             =   1950
         Width           =   855
      End
      Begin VB.Label lRZona 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Zona"
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
         Height          =   315
         Left            =   4320
         TabIndex        =   35
         Top             =   420
         Width           =   2625
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Zona:"
         Height          =   195
         Left            =   3840
         TabIndex        =   34
         Top             =   460
         Width           =   495
      End
      Begin VB.Label lID 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   840
         TabIndex        =   33
         Top             =   60
         Width           =   1005
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Id:"
         Height          =   195
         Left            =   60
         TabIndex        =   32
         Top             =   60
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha:"
         Height          =   255
         Left            =   60
         TabIndex        =   0
         Top             =   460
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Asignado:"
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Impreso:"
         Height          =   255
         Left            =   4800
         TabIndex        =   31
         Top             =   60
         Width           =   735
      End
      Begin VB.Label lRImpreso 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/09/00 23:55"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5520
         TabIndex        =   30
         Top             =   60
         Width           =   1425
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "&Costo:"
         Height          =   255
         Left            =   60
         TabIndex        =   7
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "F/&Pago:"
         Height          =   255
         Left            =   2820
         TabIndex        =   10
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "&Flete:"
         Height          =   255
         Left            =   2820
         TabIndex        =   5
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "&Liquidar:"
         Height          =   255
         Left            =   4980
         TabIndex        =   12
         Top             =   1200
         Width           =   615
      End
   End
   Begin VB.Frame frmDatos 
      Caption         =   "Datos Servicio"
      ForeColor       =   &H00800000&
      Height          =   1275
      Left            =   60
      TabIndex        =   20
      Top             =   480
      Width           =   7095
      Begin VB.TextBox tDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   900
         Width           =   5980
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   900
         Width           =   735
      End
      Begin VB.Label lTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2520
         TabIndex        =   27
         Top             =   600
         Width           =   4485
      End
      Begin VB.Label lIdProducto 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   26
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Producto:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   795
      End
      Begin VB.Label lPTipo 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   2040
         TabIndex        =   24
         Top             =   600
         Width           =   435
      End
      Begin VB.Label lEstado 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2520
         TabIndex        =   23
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Servicio:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lServicio 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   1020
         TabIndex        =   21
         Top             =   240
         Width           =   1185
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   4665
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   12726
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.TabStrip tbEntrega 
      Height          =   435
      Left            =   0
      TabIndex        =   28
      Top             =   1800
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      TabFixedWidth   =   3350
      TabFixedHeight  =   441
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6060
      Top             =   -60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEntrega.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEntrega.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEntrega.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEntrega.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEntrega.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEntrega.frx":0BA4
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
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "&Del formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sNuevo As Boolean, sModificar As Boolean
Dim gServicio As Long, gCliente As Long

Dim bHacerNota As Boolean, bHacerAnulacion As Boolean
Dim gDocumentoQFactura As Long, gDocumentoService As Long
Dim iTam As Long
Dim aNotaAImprimir As Long
Dim gSucesoUsr As Long, gSucesoDef As String

Private Sub loc_DefinoPrecioFlete(ByVal iTipoFlete As Long, ByVal iZona As Long, ByRef cValorFlete As Currency, ByRef cLiquidar As Currency)
On Error GoTo errDPF
   
    
    Cons = "SELECT Top 1 PFLPrecioPpal, PFLCostoPpal FROM PrecioFlete, GrupoZonaZona" & _
        " WHERE PFlTipoFlete = " & iTipoFlete & " AND GZZZona = " & iZona & _
        " AND  PFlGrupoZona = GZZGrupo"
    
    If iTam > 0 Then Cons = Cons & " And PFlTamañoArt = " & iTam
    
    Cons = Cons & " Order by PFLPrecioPpal"
    
    'PFlTamañoArt
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux("PFLPrecioPpal")) Then cValorFlete = RsAux("PFLPrecioPpal")
        If Not IsNull(RsAux("PFLCostoPpal")) Then cLiquidar = RsAux("PFLCostoPpal")
    Else
        If iTam > 0 Then
            'busco el menor precio.
            RsAux.Close
            Cons = "SELECT Top 1 PFLPrecioPpal, PFLCostoPpal FROM PrecioFlete, GrupoZonaZona" & _
                " WHERE PFlTipoFlete = " & iTipoFlete & " AND GZZZona = " & iZona & _
                " AND  PFlGrupoZona = GZZGrupo Order by PFLPrecioPpal"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                If Not IsNull(RsAux("PFLPrecioPpal")) Then cValorFlete = RsAux("PFLPrecioPpal")
                If Not IsNull(RsAux("PFLCostoPpal")) Then cLiquidar = RsAux("PFLCostoPpal")
            End If
        End If
    End If
    RsAux.Close
    Exit Sub
    
errDPF:
    clsGeneral.OcurrioError "Error al buscar el costo del flete.", Err.Description, "Definir precio del flete"
End Sub

Public Property Get prmServicio() As Long
    prmServicio = gServicio
End Property
Public Property Let prmServicio(Codigo As Long)
    gServicio = Codigo
End Property

Private Sub bFactura_Click()
    EjecutarApp App.Path & "\Detalle de factura", CStr(bFactura.Tag)
End Sub

Private Sub cRAsignado_GotFocus()
    If cRAsignado.Enabled Then With cRAsignado: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub cRAsignado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cRFlete
End Sub

Private Sub cRFlete_Change()
    If cRFlete.ListIndex > -1 Then BuscoValorFlete
End Sub

Private Sub cRFlete_Click()
    If cRFlete.ListIndex > -1 Then BuscoValorFlete
End Sub

Private Sub cRFlete_GotFocus()
    If cRFlete.Enabled Then With cRFlete: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub cRFlete_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tRCosto
End Sub

Private Sub cRFPago_GotFocus()
    If cRFPago.Enabled Then With cRFPago: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub cRFPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tRLiquidar
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    ObtengoSeteoForm Me
    If Me.Height <> 5610 Then Me.Height = 5610
    
    sNuevo = False: sModificar = False
    CargoCombos
    LimpioFicha Cabezal:=True
    
    DeshabilitoIngreso
    If gServicio <> 0 Then
        CargoDatosServicio
        CargoEntregas
        If tbEntrega.Tabs.Count = 0 Then AccionPrimeraEntrega
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    tbEntrega.Top = frmDatos.Top + frmDatos.Height + 100
    tbEntrega.Left = frmDatos.Left
    tbEntrega.Width = Me.ScaleWidth - tbEntrega.Left - 60
    tbEntrega.Height = Me.ScaleHeight - tbEntrega.Top - 300
    
    Dim bClear As Boolean: bClear = False
    If tbEntrega.Tabs.Count = 0 Then tbEntrega.Tabs.Add: bClear = True
    picRetiro.Top = tbEntrega.ClientTop: picRetiro.Left = tbEntrega.ClientLeft
    picRetiro.Width = tbEntrega.ClientWidth: picRetiro.Height = tbEntrega.ClientHeight
    picRetiro.BorderStyle = vbBSNone
    If bClear Then tbEntrega.Tabs.Clear
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    GuardoSeteoForm Me

End Sub

Private Sub Label3_Click()
    Foco tRFecha
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

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Private Sub AccionNuevo()
    
    On Error GoTo errNuevo
    bHacerNota = False: bHacerAnulacion = False
        
    If tbEntrega.Tabs.Count > 0 Then
        Dim aTexto As String
        aTexto = "Al realizar una nueva entrega, la anterior quedará sin efecto." & Chr(vbKeyReturn) & _
                    "Está seguro de hacer una nueva." & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                    "(*) Recuerde que al técnico/camionero se le liquidarán " & Trim(cRMoneda.Text) & " "
        If Trim(tRLiquidar.Text) = "" Then aTexto = aTexto & "0.00" Else aTexto = aTexto & Trim(tRLiquidar.Text)
        If MsgBox(aTexto, vbQuestion + vbYesNo + vbDefaultButton2, "Nueva Entrega") = vbNo Then Exit Sub
        
        If Not ValidoDocumento(Val(bFactura.Tag), bHacerNota, bHacerAnulacion, gDocumentoQFactura, gDocumentoService) Then Exit Sub
        
    End If
    sNuevo = True
    
    Botones False, False, False, True, True, Toolbar1, Me
    tbEntrega.Tabs.Add pvcaption:="Entrega (nuevo)"
    tbEntrega.Tabs(tbEntrega.Tabs.Count).Selected = True
    
    LimpioFicha ParaNuevo:=True
    HabilitoIngreso
    Foco tRFecha
    Exit Sub

errNuevo:
    clsGeneral.OcurrioError "Ocurrió un error al realizar una nueva entrega.", Err.Description
End Sub

Private Sub AccionPrimeraEntrega()
    
    Screen.MousePointer = 11
    On Error GoTo errNuevo
    sNuevo = True
    Botones False, False, False, True, True, Toolbar1, Me
    tbEntrega.Tabs.Add pvcaption:="Entrega (nuevo)"
    tbEntrega.Tabs(tbEntrega.Tabs.Count).Selected = True
    
    BuscoCodigoEnCombo cRMoneda, CLng(paMonedaPesos)
    
    If paTipoFleteVentaTelefonica = 0 Then Exit Sub
    'Cargo los controles con los datos por defecto----------------------------------------------------------------------------------------
    'Por defecto pongo que factura camión
    BuscoCodigoEnCombo cRFPago, FacturaServicio.Camion: tRLiquidar.Text = "0.00"

    Cons = "Select * From TipoFlete Where TFlCodigo = " & paTipoFleteVentaTelefonica
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!TFlFormaPago) Then
            'Factura Camión = 3
            If RsAux!TFlFormaPago = 3 Then
                BuscoCodigoEnCombo cRFPago, FacturaServicio.Camion
            Else
                BuscoCodigoEnCombo cRFPago, RsAux!TFlFormaPago
            End If
        End If
        
        If Not IsNull(RsAux!TFlAgenda) Then
            If Not IsNull(RsAux!TFlFechaAgeHab) Then
                tRFecha.Text = fnc_FindFirstDay(RsAux!TFlFechaAgeHab, RsAux!TFlAgenda, RsAux!TFlAgendaHabilitada)
            Else
                tRFecha.Text = fnc_FindFirstDay(Date, RsAux!TFlAgenda, RsAux!TFlAgenda)
            End If
        End If
        
    End If
    RsAux.Close
    
    If paCamionRetiroVisita > 0 Then
        BuscoCodigoEnCombo cRAsignado, paCamionRetiroVisita
        
        Dim RsF As rdoResultset
        Cons = "SELECT CTFTipoFlete, TFLFormaPago FROM CamionFlete INNER JOIN TipoFlete ON TFLCodigo = CTFTipoFlete WHERE CTFCamion = " & paCamionRetiroVisita
        Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsF.EOF Then
            BuscoCodigoEnCombo cRFPago, RsF!TFlFormaPago
            BuscoCodigoEnCombo cRFlete, RsF(0)

        Else
            BuscoCodigoEnCombo cRFlete, paTipoFleteVentaTelefonica
            
        End If
        
    Else
        Cons = "Select * From CamionFlete, CamionZona " _
            & " Where CTFTipoFlete = " & paTipoFleteVentaTelefonica _
            & " And CZoZona = " & Val(lRZona.Tag) _
            & " And CTFCamion = CZoCamion Order by CZoPrioridad "
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux.EOF Then BuscoCodigoEnCombo cRAsignado, RsAux!CZoCamion
        RsAux.Close
        
        BuscoCodigoEnCombo cRFlete, paTipoFleteVentaTelefonica
    End If
    
    
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    HabilitoIngreso
    Foco tRFecha
    Screen.MousePointer = 0
    Exit Sub

errNuevo:
    clsGeneral.OcurrioError "Ocurrió un error al realizar una nueva entrega.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub AccionModificar()
    
    On Error Resume Next
    sModificar = True
    bHacerNota = False: bHacerAnulacion = False
    
    HabilitoIngreso
    Botones False, False, False, True, True, Toolbar1, Me
    tUsuario.Text = ""
    If tRFecha.Enabled Then Foco tRFecha Else Foco tRLiquidar
        
End Sub

Private Sub AccionGrabar()

    If Not ValidoCampos Then Exit Sub
    
    If MsgBox("Confirma almacenar la información ingresada", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    
    If bHacerNota Or bHacerAnulacion Then
        Screen.MousePointer = 11
        Dim objSuceso As New clsSuceso
        objSuceso.ActivoFormulario paCodigoDeUsuario, "Anulación de Documentos en Servicio", cBase
        Me.Refresh
        gSucesoUsr = objSuceso.RetornoValor(Usuario:=True)
        gSucesoDef = objSuceso.RetornoValor(Defensa:=True)
        Set objSuceso = Nothing
        If gSucesoUsr = 0 Then Screen.MousePointer = 0: Exit Sub 'Abortó el ingreso del suceso
    End If
    
   aNotaAImprimir = 0
    
    On Error GoTo errorBT
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
    On Error GoTo errorET
    
    If sNuevo Then
        Cons = "Select * from ServicioVisita Where VisCodigo = 0"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.AddNew
        CargoCamposBD
        RsAux.Update: RsAux.Close
        
        If Val(tbEntrega.Tag) <> 0 Then
            Cons = "Update ServicioVisita Set VisSinEfecto = 1 Where VisCodigo = " & Val(tbEntrega.Tag)
            cBase.Execute Cons
        End If
        
        Cons = "Update Servicio Set SerEstadoServicio = " & EstadoS.Entrega & " Where SerCodigo = " & gServicio
        cBase.Execute Cons
        
        If bHacerNota Or bHacerAnulacion Then
            aNotaAImprimir = ProcesoDocumentoFacturado(gDocumentoQFactura, bHacerNota, bHacerAnulacion, gServicio, gSucesoUsr, gSucesoDef, gDocumentoService)
        End If
        
    Else                                    'Modificar----
           
        Cons = "Select * from ServicioVisita Where VisCodigo = " & Val(tbEntrega.Tabs(tbEntrega.SelectedItem.Index).Tag)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.Edit
        CargoCamposBD
        RsAux.Update: RsAux.Close
    End If
    cBase.CommitTrans    'FIN TRANSACCION------------------------------------------
    
    sNuevo = False: sModificar = False
    DeshabilitoIngreso
    LimpioFicha
    CargoEntregas
    
    If bHacerNota And aNotaAImprimir <> 0 Then AccionEmitirNota aNotaAImprimir
    
    Screen.MousePointer = 0
    Exit Sub
    
errorBT:
    clsGeneral.OcurrioError "Ha ocurrido un error al realizar la operación.", Err.Description
    Screen.MousePointer = 0: Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ha ocurrido un error al realizar la operación.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionEliminar()
    
    Dim aTexto As String
    aTexto = "Confirma eliminar la entrega seleccionada." & Chr(vbKeyReturn) & _
                 "Si la elimina, la entrega quedará sin efecto." & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                 "(*) Recuerde que al técnico/camionero se le liquidarán " & Trim(cRMoneda.Text) & " "
    If Trim(tRLiquidar.Text) = "" Then aTexto = aTexto & "0.00" Else aTexto = aTexto & Trim(tRLiquidar.Text)

    If MsgBox(aTexto, vbQuestion + vbYesNo + vbDefaultButton2, "ELIMINAR") = vbNo Then Screen.MousePointer = 0: Exit Sub
    
    If Not ValidoDocumento(Val(bFactura.Tag), bHacerNota, bHacerAnulacion, gDocumentoQFactura, gDocumentoService) Then Exit Sub
    
    Screen.MousePointer = 11
    aNotaAImprimir = 0
    
    If bHacerNota Or bHacerAnulacion Then       'Suceso
        Screen.MousePointer = 11
        Dim objSuceso As New clsSuceso
        objSuceso.ActivoFormulario paCodigoDeUsuario, "Anulación de Documentos en Servicio", cBase
        Me.Refresh
        gSucesoUsr = objSuceso.RetornoValor(Usuario:=True)
        gSucesoDef = objSuceso.RetornoValor(Defensa:=True)
        Set objSuceso = Nothing
        If gSucesoUsr = 0 Then Screen.MousePointer = 0: Exit Sub 'Abortó el ingreso del suceso
    End If

    On Error GoTo errorBT
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
    On Error GoTo errorET
    
    Cons = "Select * from ServicioVisita Where VisCodigo = " & Val(tbEntrega.Tabs(tbEntrega.SelectedItem.Index).Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.Edit
    RsAux!VisSinEfecto = True
    RsAux.Update: RsAux.Close
    
    If bHacerNota Or bHacerAnulacion Then
        aNotaAImprimir = ProcesoDocumentoFacturado(gDocumentoQFactura, bHacerNota, bHacerAnulacion, gServicio, gSucesoUsr, gSucesoDef, gDocumentoService)
    End If
        
    cBase.CommitTrans    'FIN TRANSACCION------------------------------------------
    
    LimpioFicha
    CargoEntregas
    
    If bHacerNota And aNotaAImprimir <> 0 Then AccionEmitirNota aNotaAImprimir
    Screen.MousePointer = 0
    Exit Sub
    
errorBT:
    clsGeneral.OcurrioError "Ha ocurrido un error al realizar la operación.", Err.Description
    Screen.MousePointer = 0: Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ha ocurrido un error al realizar la operación.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionEmitirNota(idNota As Long)
    
    On Error GoTo errImprimir
    Screen.MousePointer = 11
    Status.Panels(1).Text = "Abriendo motor de impresión..."
    If crAbroEngine = 0 Then     'Abro el Engine del Crystal
        Screen.MousePointer = 11
        clsGeneral.OcurrioError Trim(crMsgErr), Err.Description
         Status.Panels(1).Text = "": Screen.MousePointer = 0
        Exit Sub
    End If
    
    Status.Panels(1).Text = "Imprimiendo Documento..."
    ImprimoNota idNota, gDocumentoQFactura, gCliente
    
     Status.Panels(1).Text = "Cerrando motor de impresión..."
    crCierroEngine      'Cierro el Engine del Crystal
    Status.Panels(1).Text = ""
    Screen.MousePointer = 0
    Exit Sub
    
errImprimir:
    clsGeneral.OcurrioError "Ocurrió un error al imprimir la nota.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionCancelar()

    On Error Resume Next
    DeshabilitoIngreso
    LimpioFicha
    sNuevo = False: sModificar = False
    CargoEntregas
    
End Sub

Private Sub tbEntrega_Click()
    
    If tbEntrega.Tabs.Count = 0 Or sNuevo Then Exit Sub
    CargoDatosEntrega tbEntrega.Tabs(tbEntrega.SelectedItem.Index).Tag
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        Case "salir": Unload Me
    End Select

End Sub

Private Function ValidoCampos() As Boolean

    ValidoCampos = False
    
    If Not IsDate(tRFecha.Text) Then
        MsgBox "La fecha ingresada no es correcta.", vbExclamation, "ATENCIÓN": Foco tRFecha: Exit Function
    Else
        If tRFecha.Enabled And Trim(lRImpreso.Caption) = "" _
            And CDate(tRFecha.Text) < CDate(Format(gFechaServidor, "dd/mm/yyyy")) Then
            If MsgBox("La fecha de entrega es menor a la fecha de hoy. Desea continuar.", vbExclamation + vbYesNo + vbDefaultButton2, "ATENCIÓN") = vbNo Then Exit Function
        End If
    End If
            
    If Trim(cRHora.Text) = "" And cRHora.Enabled Then
        MsgBox "Debe ingresar la hora para realizar la entrega.", vbExclamation, "Faltan Datos"
        Foco cRHora: Exit Function
    End If
    
    If cRAsignado.ListIndex = -1 And cRAsignado.Enabled Then
        MsgBox "Seleccione a quien se va a asignar la entrega del producto.", vbExclamation, "Faltan Datos"
        Foco cRAsignado: Exit Function
    End If
    
    If cRFlete.ListIndex = -1 And cRFlete.Enabled Then
        MsgBox "Seleccione el tipo de flete para la entrega.", vbExclamation, "Faltan Datos"
        Foco cRFlete: Exit Function
    End If
       
    If cRMoneda.ListIndex = -1 And cRMoneda.Enabled Then
        MsgBox "Seleccione la moneda para ingresar el costo de la entrega.", vbExclamation, "Faltan Datos"
        Foco cRMoneda: Exit Function
    End If
    If Not IsNumeric(tRCosto.Text) And tRCosto.Enabled Then
        MsgBox "Ingrese el costo de la entrega.", vbExclamation, "Faltan Datos"
        Foco tRCosto: Exit Function
    End If
    
    If cRFPago.ListIndex = -1 And cRFPago.Enabled Then
        MsgBox "Seleccione la forma de pago de la entrega.", vbExclamation, "Faltan Datos"
        Foco cRFPago: Exit Function
    End If
    If Not IsNumeric(tRLiquidar.Text) And tRLiquidar.Enabled Then
        MsgBox "Ingrese el importe que se va a liquidar al camionero.", vbExclamation, "Faltan Datos"
        Foco tRLiquidar: Exit Function
    End If
    
    If Val(tUsuario.Tag) = 0 Then
        MsgBox "Ingrese su dígito de usuario para grabar.", vbExclamation, "Faltan Datos"
        Foco tUsuario: Exit Function
    End If
    
    'Veo si tiene retiro y le doy aviso que valide si no se cobro en el mismo.
    If Val(tRCosto.Text) > 0 And tRCosto.Enabled Then
        If TieneRetiroConCobro Then
            If MsgBox("El servicio tiene un retiro con cobro, está seguro de no duplicar?." & vbCrLf & "¿Esta seguro que desea cobrarle la entrega?", vbQuestion + vbYesNo + vbDefaultButton2, "ATENCIÓN") = vbNo Then
                Exit Function
            End If
        End If
        If Not ValidoLiquidarCamion Then Exit Function
    End If
    
    ValidoCampos = True
    
End Function

Private Function TieneRetiroConCobro() As Boolean
On Error GoTo errTRC
Dim rsV As rdoResultset
    TieneRetiroConCobro = False
    Cons = "Select * From ServicioVisita Where VisServicio = " & Val(lServicio.Caption) _
        & " And VisTipo = " & TipoServicio.Retiro & " And VisSinEfecto = 0 And VisCosto > 0"
    Set rsV = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsV.EOF Then TieneRetiroConCobro = True
    rsV.Close
errTRC:
End Function

Private Function ValidoLiquidarCamion() As Boolean
On Error GoTo errTRC
Dim rsV As rdoResultset
Dim cRet As Currency, cTot As Currency

    ValidoLiquidarCamion = True
    If Val(tRLiquidar.Text) = 0 Then Exit Function
    
    Cons = "Select * From ServicioVisita Where VisServicio = " & Val(lServicio.Caption) _
        & " And VisTipo = " & TipoServicio.Retiro & " And VisSinEfecto = 0 And VisCosto > 0"
    Set rsV = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsV.EOF Then
        If Not IsNull(rsV!VisLiquidarAlCamion) Then cRet = rsV!VisLiquidarAlCamion
    End If
    rsV.Close
    
    Dim cVF As Currency, cLi As Currency
    loc_DefinoPrecioFlete cRFlete.ItemData(cRFlete.ListIndex), lRZona.Tag, cVF, cLi
    cTot = Format(cVF * paCoefFleteRetiro, FormatoMonedaP)
    
    If CCur(tRLiquidar.Text) + cRet > cTot Then
        If MsgBox("No le estará liquidando de más al camionero" & vbCrLf & "¿Desea almacenar el importe ingresado?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
            ValidoLiquidarCamion = False
        End If
    End If
    
errTRC:
End Function


Private Sub DeshabilitoIngreso()
    
    tbEntrega.Enabled = True
    
    tRFecha.Enabled = False: tRFecha.BackColor = Inactivo
    cRHora.Enabled = False: cRHora.BackColor = Inactivo
    
    cRAsignado.Enabled = False: cRAsignado.BackColor = Colores.Inactivo
    cRFlete.Enabled = False: cRFlete.BackColor = Colores.Inactivo
    cRMoneda.Enabled = False: cRMoneda.BackColor = Colores.Inactivo
    tRCosto.Enabled = False: tRCosto.BackColor = Inactivo
    cRFPago.Enabled = False: cRFPago.BackColor = Colores.Inactivo
    tRLiquidar.Enabled = False: tRLiquidar.BackColor = Inactivo
    
    tRComentario.Enabled = False: tRComentario.BackColor = Inactivo
    tUsuario.Enabled = False: tUsuario.BackColor = Inactivo
        
End Sub

Private Sub HabilitoIngreso()

    tbEntrega.Enabled = False
    
    If lRSinEfecto.Visible Then
        tRLiquidar.Enabled = True: tRLiquidar.BackColor = Blanco
        tRComentario.Enabled = True: tRComentario.BackColor = Blanco
        tUsuario.Enabled = True: tUsuario.BackColor = Blanco
        Exit Sub
    End If
    
    If sNuevo Or (sModificar And Trim(lRImpreso.Caption) = "") Then
        tRFecha.Enabled = True: tRFecha.BackColor = Blanco
        cRHora.Enabled = True: cRHora.BackColor = Blanco
        
        cRAsignado.Enabled = True: cRAsignado.BackColor = Colores.Blanco
        cRFlete.Enabled = True: cRFlete.BackColor = Colores.Blanco
        cRMoneda.Enabled = True: cRMoneda.BackColor = Colores.Blanco
        tRCosto.Enabled = True: tRCosto.BackColor = Blanco
        cRFPago.Enabled = True: cRFPago.BackColor = Colores.Blanco
    ElseIf (sModificar) Then
        If miConexion.AccesoAlMenu("DecisiónReparto") Then
            tRFecha.Enabled = True: tRFecha.BackColor = Blanco
            cRHora.Enabled = True: cRHora.BackColor = Blanco
            cRAsignado.Enabled = True: cRAsignado.BackColor = Colores.Blanco
        End If
        
    End If
    
    tRLiquidar.Enabled = True: tRLiquidar.BackColor = Blanco
    tRComentario.Enabled = True: tRComentario.BackColor = Blanco
    tUsuario.Enabled = True: tUsuario.BackColor = Blanco
    
End Sub

Private Sub CargoCamposBD()
    
    RsAux!VisServicio = CLng(lServicio.Caption)
    RsAux!VisTipo = TipoServicio.Entrega

    If sNuevo Then RsAux!VisSinEfecto = False
        
    RsAux!VisFecha = Format(tRFecha.Text, sqlFormatoF)
    If cRHora.ListIndex = -1 Then
        RsAux!VisHorario = Trim(cRHora.Text)
    Else
        Dim rsHor As rdoResultset
        Cons = "Select * from CodigoTexto Where Codigo = " & cRHora.ItemData(cRHora.ListIndex)
        Set rsHor = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsHor.EOF Then
            RsAux!VisHorario = Format(rsHor!Clase, "0000") & "-" & Format(rsHor!Puntaje, "0000")
        Else
            RsAux!VisHorario = Trim(cRHora.Text)
        End If
        rsHor.Close
    End If
    RsAux!VisCamion = cRAsignado.ItemData(cRAsignado.ListIndex)
    RsAux!VisZona = Val(lRZona.Tag)
    
    RsAux!VisMoneda = cRMoneda.ItemData(cRMoneda.ListIndex)
    RsAux!VisCosto = CCur(tRCosto.Text)
    RsAux!VisFormaPago = cRFPago.ItemData(cRFPago.ListIndex)
     
    RsAux!VisTipoFlete = cRFlete.ItemData(cRFlete.ListIndex)
    RsAux!VisLiquidarAlCamion = CCur(tRLiquidar.Text)
    
    If Trim(tRComentario.Text) <> "" Then RsAux!VisComentario = Trim(tRComentario.Text) Else RsAux!VisComentario = Null
    
    RsAux!VisFModificacion = Format(gFechaServidor, sqlFormatoFH)
    RsAux!VisUsuario = Val(tUsuario.Tag)
    
End Sub

Private Sub LimpioFicha(Optional Cabezal As Boolean = False, Optional ParaNuevo As Boolean = False)

    If Cabezal Then
        lServicio.Caption = "": lEstado.Caption = ""
        lIdProducto.Caption = "": lTipo.Caption = ""
        tDireccion.Text = "": tDireccion.Tag = 0
    End If

    lID.Caption = "": lRImpreso.Caption = ""
    tRFecha.Text = ""
    
    If Not ParaNuevo Then
        cRHora.Text = "": lRZona.Caption = "": lRZona.Tag = 0
        
        cRAsignado.Text = "": cRFlete.Text = ""
        cRMoneda.Text = "": tRCosto.Text = "": cRFPago.Text = "": tRLiquidar.Text = ""
        
        tRComentario.Text = ""
    End If
    
    tUsuario.Text = "": lRModificado.Caption = ""
    lRSinEfecto.Visible = False
    chLiquidada.Value = vbUnchecked
    bFactura.Tag = 0: bFactura.Enabled = False
        
End Sub

Private Sub CargoDatosServicio()

    On Error GoTo errCargar
    iTam = 0
    Cons = "Select * from Servicio, Producto, Articulo " & _
               " Where SerCodigo = " & gServicio & _
               " And SerProducto = ProCodigo " & _
               " And ProArticulo = ArtID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        If Not IsNull(RsAux("ArtTamaño")) Then iTam = RsAux("ArtTamaño")
        lServicio.Caption = " " & RsAux!SerCodigo
        lEstado.Caption = UCase(EstadoServicio(RsAux!SerEstadoServicio))
        lEstado.Tag = RsAux!SerEstadoServicio
        
        lIdProducto.Caption = " " & Format(RsAux!ProCodigo, "000")
        lTipo.Caption = " " & Trim(RsAux!ArtNombre)
        
        If Not IsNull(RsAux!ProCliente) Then gCliente = RsAux!ProCliente Else gCliente = 0
        If Not IsNull(RsAux!SerDocumento) Then gDocumentoService = RsAux!SerDocumento Else gDocumentoService = 0
        
        If Not IsNull(RsAux!ProDireccion) Then
            tDireccion.Text = " " & clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!ProDireccion, True, True, True)
            tDireccion.Tag = RsAux!ProDireccion
        Else
            Dim rsCli As rdoResultset
            Cons = "Select * from Cliente Where CliCodigo = " & RsAux!ProCliente
            Set rsCli = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rsCli.EOF Then
                If Not IsNull(rsCli!CliDireccion) Then
                    tDireccion.Text = " " & clsGeneral.ArmoDireccionEnTexto(cBase, rsCli!CliDireccion, True, True, True)
                    tDireccion.Tag = rsCli!CliDireccion
                End If
            End If
            rsCli.Close
        End If
        
    End If
    RsAux.Close
    
    If Val(tDireccion.Tag) <> 0 Then
        Dim auxZona As Long
        auxZona = BuscoZonaDireccion(CLng(tDireccion.Tag))
        Cons = "Select * from Zona where ZonCodigo = " & auxZona
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            lRZona.Caption = " " & Trim(RsAux!ZonNombre)
            lRZona.Tag = auxZona
        End If
        RsAux.Close
    End If
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del servicio.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoEntregas()

    On Error GoTo errCargar
    Dim aIdEntrega As Long: aIdEntrega = 0
    tbEntrega.Tabs.Clear
    'Armo los Tabs con todos las entregas y cargo el mayor
    
    Cons = "Select * from ServicioVisita " & _
               " Where VisServicio = " & gServicio & _
               " And VisTipo = " & TipoServicio.Entrega & _
               " Order by VisCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        tbEntrega.Tabs.Add pvcaption:="Entrega (" & RsAux!VisCodigo & ")"
        tbEntrega.Tabs.Item(tbEntrega.Tabs.Count).Tag = RsAux!VisCodigo
        aIdEntrega = RsAux!VisCodigo
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    tbEntrega.Tag = aIdEntrega            'el maximo entrega para el servicio
    
    For I = 1 To tbEntrega.Tabs.Count
        If tbEntrega.Tabs(I).Tag = aIdEntrega Then tbEntrega.Tabs(I).Selected = True
    Next
            
    If tbEntrega.Tabs.Count = 0 Then Botones True, False, False, False, False, Toolbar1, Me
    
    Exit Sub
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos de las entregas.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoDatosEntrega(idEntrega As Long)
    
Dim auxZona As Long

    On Error GoTo errCargar
    auxZona = 0
    Screen.MousePointer = 11
    LimpioFicha
    
    Cons = "Select * from ServicioVisita Left Outer Join Zona On VisZona = ZonCodigo" & _
               " Where VisCodigo = " & idEntrega
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        lID.Caption = " " & idEntrega
        
        tRFecha.Text = Format(RsAux!VisFecha, "dd/mm/yyyy")
        cRHora.Text = Trim(RsAux!VisHorario)
        If Not IsNull(RsAux!ZonNombre) Then
            lRZona.Caption = " " & Trim(RsAux!ZonNombre)
            lRZona.Tag = RsAux!ZonCodigo
        End If
        
        If Not IsNull(RsAux!VisFImpresion) Then lRImpreso.Caption = Format(RsAux!VisFImpresion, "dd/mm/yy hh:mm")
        BuscoCodigoEnCombo cRAsignado, RsAux!VisCamion
        BuscoCodigoEnCombo cRMoneda, RsAux!VisMoneda
        tRCosto.Text = Format(RsAux!VisCosto, FormatoMonedaP)
        BuscoCodigoEnCombo cRFPago, RsAux!VisFormaPago
        If Not IsNull(RsAux!VisTipoFlete) Then BuscoCodigoEnCombo cRFlete, RsAux!VisTipoFlete
        If Not IsNull(RsAux!VisComentario) Then tRComentario.Text = Trim(RsAux!VisComentario)
        
        If Not IsNull(RsAux!VisLiquidarAlCamion) Then tRLiquidar.Text = Format(RsAux!VisLiquidarAlCamion, FormatoMonedaP)
        
        If Not IsNull(RsAux!VisFModificacion) Then lRModificado.Caption = " " & Format(RsAux!VisFModificacion, "dd/mm/yy hh:mm")
        If Not IsNull(RsAux!VisUsuario) Then tUsuario.Text = BuscoUsuario(RsAux!VisUsuario, Identificacion:=True)
        
        If RsAux!VisSinEfecto Then lRSinEfecto.Visible = True Else lRSinEfecto.Visible = False
        If Not IsNull(RsAux!VisLiquidada) Then chLiquidada.Value = vbChecked Else chLiquidada.Value = vbUnchecked
        
        If Not IsNull(RsAux!VisDocumento) Then
            bFactura.Enabled = True: bFactura.Tag = RsAux!VisDocumento
        Else
            bFactura.Enabled = False: bFactura.Tag = 0
        End If
        
        'Comparo la Zona para updatear---------------------------------------------------------------------------
        If Not lRSinEfecto.Visible Then
            auxZona = BuscoZonaDireccion(CLng(tDireccion.Tag))
            If auxZona <> lRZona.Tag And auxZona <> 0 Then
                RsAux.Edit
                RsAux!VisZona = auxZona
                RsAux.Update
            End If
        End If
        '----------------------------------------------------------------------------------------------------------------
    End If
    RsAux.Close
    
    If auxZona <> Val(lRZona.Tag) And auxZona <> 0 Then
        Cons = "Select * from Zona Where ZonCodigo = " & auxZona
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            lRZona.Caption = " " & Trim(RsAux!ZonNombre)
            lRZona.Tag = RsAux!ZonCodigo
        End If
        RsAux.Close
    End If
    
    Screen.MousePointer = 0
    
    If lRSinEfecto.Visible Then
        If Val(tbEntrega.Tag) <> idEntrega Then Botones False, True, False, False, False, Toolbar1, Me Else Botones True, True, False, False, False, Toolbar1, Me
        Exit Sub
    End If
    If Val(tbEntrega.Tag) <> idEntrega Then Botones False, False, False, False, False, Toolbar1, Me: Exit Sub
    If Val(lEstado.Tag) <> EstadoS.Entrega Then Botones False, True, False, False, False, Toolbar1, Me: Exit Sub
        
    If Trim(lRImpreso.Caption) = "" Then
        Botones False, True, True, False, False, Toolbar1, Me
    Else
        Botones True, True, True, False, False, Toolbar1, Me
    End If
    Exit Sub

errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos de la entrega.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoCombos()
    
    Cons = "Select * from Moneda Where MonFactura = 1 Order by MonSigno"
    CargoCombo Cons, cRMoneda
    
    Cons = "Select * from Camion order by CamNombre"
    CargoCombo Cons, cRAsignado
    
    cRFPago.AddItem TipoFacturaServicio(FacturaServicio.Camion): cRFPago.ItemData(cRFPago.NewIndex) = FacturaServicio.Camion
    cRFPago.AddItem TipoFacturaServicio(FacturaServicio.CGSA): cRFPago.ItemData(cRFPago.NewIndex) = FacturaServicio.CGSA
    cRFPago.AddItem TipoFacturaServicio(FacturaServicio.SinFactura): cRFPago.ItemData(cRFPago.NewIndex) = FacturaServicio.SinFactura
    
    Cons = "Select * from TipoFlete order by TFlNombreCorto"
    CargoCombo Cons, cRFlete
    
    Cons = "Select Codigo, Texto from CodigoTexto Where Tipo = 15 Order by texto"
    CargoCombo Cons, cRHora
    
End Sub

Private Sub tRComentario_GotFocus()
    With tRComentario: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tRComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tUsuario
End Sub

Private Sub tRCosto_GotFocus()
    With tRCosto: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tRCosto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cRFPago
End Sub

Private Sub tRFecha_GotFocus()
    With tRFecha: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tRFecha_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        tRFecha.Text = Format(tRFecha.Text, "dd/mm/yyyy")
        If IsDate(tRFecha.Text) Then
            
            If CDate(tRFecha.Text) < CDate(Format(gFechaServidor, "dd/mm/yyyy")) Then
                MsgBox "La fecha de entrega no debe ser menor a la fecha de hoy.", vbExclamation, "ATENCIÓN"
            Else
                Foco cRHora
            End If
        End If
    End If
    
End Sub

Private Sub tRFecha_LostFocus()
    If IsDate(tRFecha.Text) Then tRFecha.Text = Format(tRFecha.Text, "dd/mm/yyyy")
End Sub

Private Sub cRHora_GotFocus()
    With cRHora: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub cRHora_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(cRHora.Text) = "" Or cRHora.ListIndex <> -1 Then Foco cRAsignado: Exit Sub
        cRHora.Text = ValidoRangoHorario(cRHora.Text)
        If Trim(cRHora.Text) <> "" Then Foco cRAsignado
    End If
    
End Sub

Private Sub tRLiquidar_GotFocus()
    With tRLiquidar: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tRLiquidar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Not IsNumeric(tRLiquidar.Text) Then tRLiquidar.Text = 0
        If IsNumeric(tRLiquidar.Text) Then tRLiquidar.Text = Format(tRLiquidar.Text, FormatoMonedaP)
        Foco tRComentario
    End If
End Sub

Private Sub tUsuario_Change()
    tUsuario.Tag = 0
End Sub

Private Sub tUsuario_GotFocus()
    With tUsuario: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And Val(tUsuario.Tag) <> 0 Then AccionGrabar: Exit Sub
    
    If KeyAscii = vbKeyReturn Then
        If Not IsNumeric(tUsuario.Text) Then Exit Sub
        Dim aId As Long
        aId = BuscoUsuarioDigito(CLng(tUsuario.Text), Codigo:=True)
        tUsuario.Text = BuscoUsuario(aId, Identificacion:=True)
        tUsuario.Tag = aId
        If Val(tUsuario.Tag) <> 0 Then AccionGrabar
    End If
    
    
End Sub

Private Sub BuscoValorFlete()
    
    On Error GoTo errValor
    If cRMoneda.ListIndex = -1 Or cRFlete.ListIndex = -1 Then Exit Sub
    If Not sNuevo And Not sModificar Then Exit Sub
    
    Dim cVF As Currency, cLi As Currency, cVFReal As Currency, cLiReal As Currency
    loc_DefinoPrecioFlete cRFlete.ItemData(cRFlete.ListIndex), lRZona.Tag, cVF, cLi
    
    cVFReal = cVF
    cLiReal = cLi
    
    cVF = cVF * paCoefFleteRetiro
    cLi = cLi * paCoefFleteRetiro
    
    If TieneRetiroConCobro Then
        If MsgBox("Este servicio tiene retiro con cobro." & vbCrLf & "¿Se va a cobrar entrega?", vbQuestion + vbYesNo + vbDefaultButton2, "Posible Duplicación") = vbYes Then
            tRCosto.Text = ""
            tRCosto.Text = Format(cVF, FormatoMonedaP)
            If cRFPago.ListIndex > -1 Then
                If cLi > 0 And cRFPago.ItemData(cRFPago.ListIndex) <> FacturaServicio.Camion Then tRLiquidar.Text = Format(cLi, FormatoMonedaP) Else tRLiquidar.Text = "0.00"
            End If
        Else
            tRCosto.Text = "0.00"
            tRLiquidar.Text = "0.00"
        End If
    Else
        If paCobroEnEntrega Then
            MsgBox "Se va a cobrar retiro y entrega preste atención en el valor del flete.", vbInformation, "Retiro y entrega"
            cVF = cVFReal * 2
            cLi = cLiReal * 2
        End If
        tRCosto.Text = Format(cVF, FormatoMonedaP)
        If cRFPago.ListIndex > -1 Then
            If cLi > 0 And cRFPago.ItemData(cRFPago.ListIndex) <> FacturaServicio.Camion Then tRLiquidar.Text = Format(cLi, FormatoMonedaP) Else tRLiquidar.Text = "0.00"
        End If
    End If
    Exit Sub

errValor:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el valor del flete.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function fnc_FindFirstDay(ByVal dCierre As Date, ByVal doAgenda As Currency, ByVal doAgendaAbierta As Double) As String
Dim sMat As String
Dim iSuma As Integer
    
    fnc_FindFirstDay = Format(Date, "dd/mm/yyyy")
    If dCierre < Date Then dCierre = Date
    
    If DateDiff("d", dCierre, Date) >= 7 Then
        'Como cerro hace una semana tomo la agenda normal.
        sMat = superp_MatrizSuperposicion(doAgenda)
    Else
        sMat = superp_MatrizSuperposicion(doAgendaAbierta)
    End If
    
    If sMat <> "" Then
        iSuma = fnc_FindNextOpenDay(dCierre, sMat)
        If iSuma <> -1 Then
            fnc_FindFirstDay = Format(DateAdd("d", iSuma, dCierre), "dd/mm/yyyy")
        End If
    End If

End Function

Private Function fnc_FindNextOpenDay(dFecha As Date, strMat As String) As Integer
Dim rsHora As rdoResultset
Dim intDia As Integer, intSuma As Integer
    
    'Por las dudas que no cumpla en la semana paso la agenda normal.
    
    On Error GoTo errBDER
    
    fnc_FindNextOpenDay = -1
    
    'Consulto en base a la matriz devuelta.
    Cons = "Select * From HorarioFlete Where HFlIndice IN (" & strMat & ")"
    Set rsHora = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsHora.EOF Then
        
        'Busco el valor que coincida con el dia de hoy y ahí busco para arriba.
        intSuma = 0
        Do While intSuma < 7
            rsHora.MoveFirst
            intDia = Weekday(dFecha + intSuma)
            Do While Not rsHora.EOF
                If rsHora!HFlDiaSemana = intDia Then
                    fnc_FindNextOpenDay = intSuma
                    GoTo Encontre
                End If
                rsHora.MoveNext
            Loop
            intSuma = intSuma + 1
        Loop
        rsHora.Close
    End If

Encontre:
    rsHora.Close
    Exit Function
    
errBDER:
    clsGeneral.OcurrioError "Error al buscar el primer día disponible para el tipo de flete.", Trim(Err.Description)
End Function


