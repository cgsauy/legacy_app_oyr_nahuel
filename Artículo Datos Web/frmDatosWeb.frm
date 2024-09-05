VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{683364A1-B37D-11D1-ADC5-006008A5848C}#1.0#0"; "dhtmled.ocx"
Begin VB.Form frmDatosWeb 
   Caption         =   "Datos Web"
   ClientHeight    =   5775
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11355
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDatosWeb.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   11355
   Begin VB.PictureBox picFoto 
      Height          =   3255
      Left            =   8880
      ScaleHeight     =   3195
      ScaleWidth      =   1875
      TabIndex        =   14
      Top             =   1680
      Width           =   1935
      Begin VB.PictureBox picScroll 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2640
         ScaleHeight     =   195
         ScaleWidth      =   1395
         TabIndex        =   17
         Top             =   3240
         Width           =   1395
         Begin VB.HScrollBar hsFoto 
            Height          =   195
            LargeChange     =   100
            Left            =   0
            SmallChange     =   10
            TabIndex        =   18
            Top             =   0
            Width           =   1155
         End
      End
      Begin VB.VScrollBar vsFoto 
         Height          =   915
         LargeChange     =   100
         Left            =   4800
         TabIndex        =   16
         Top             =   0
         Width           =   195
      End
      Begin VB.Image imgFoto 
         Height          =   4200
         Left            =   60
         Top             =   60
         Width           =   2400
      End
   End
   Begin VB.TextBox tIntranet 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      MaxLength       =   20
      TabIndex        =   9
      Top             =   1260
      Width           =   3375
   End
   Begin VB.TextBox tWEB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1500
      MaxLength       =   20
      TabIndex        =   7
      Top             =   1260
      Width           =   3375
   End
   Begin VB.PictureBox picTexto 
      Height          =   2775
      Left            =   480
      ScaleHeight     =   2715
      ScaleWidth      =   7575
      TabIndex        =   12
      Top             =   2640
      Width           =   7635
      Begin VB.TextBox tTexto 
         Appearance      =   0  'Flat
         Height          =   1575
         Left            =   840
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   720
         Width           =   3315
      End
      Begin MSComctlLib.Toolbar tooTexto 
         Height          =   330
         Left            =   60
         TabIndex        =   19
         Top             =   0
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "viñeta"
               Object.ToolTipText     =   "Viñeta"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "enter"
               Object.ToolTipText     =   "Enter"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "negrita"
               Object.ToolTipText     =   "Negrita"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "subrayar"
               Object.ToolTipText     =   "Subrayado"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "glosario"
               Object.ToolTipText     =   "Glosario"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "medidas"
               Object.ToolTipText     =   "Medidas"
            EndProperty
         EndProperty
      End
      Begin DHTMLEDLibCtl.DHTMLEdit tHtml 
         Height          =   855
         Left            =   5100
         TabIndex        =   13
         Top             =   600
         Width           =   1815
         ActivateApplets =   0   'False
         ActivateActiveXControls=   0   'False
         ActivateDTCs    =   -1  'True
         ShowDetails     =   0   'False
         ShowBorders     =   0   'False
         Appearance      =   1
         Scrollbars      =   -1  'True
         ScrollbarAppearance=   1
         SourceCodePreservation=   -1  'True
         AbsoluteDropMode=   0   'False
         SnapToGrid      =   0   'False
         SnapToGridX     =   50
         SnapToGridY     =   50
         BrowseMode      =   -1  'True
         UseDivOnCarriageReturn=   0   'False
      End
   End
   Begin MSComctlLib.TabStrip tsTexto 
      Height          =   4155
      Left            =   60
      TabIndex        =   10
      Top             =   1680
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7329
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Características"
            Key             =   "caracteristicas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Datos a &Remarcar"
            Key             =   "datos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Foto"
            Key             =   "foto"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.TextBox tPrioridad 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   5760
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox tFoto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2220
      MaxLength       =   20
      TabIndex        =   5
      Top             =   900
      Width           =   2655
   End
   Begin VB.TextBox tArticulo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   900
      TabIndex        =   1
      ToolTipText     =   "Ingrese el nombre o el código de un artículo."
      Top             =   540
      Width           =   3975
   End
   Begin MSComctlLib.Toolbar tooMenu 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir del Formulario. [Ctrl+X]"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar.[Ctrl+M]"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar. [Ctrl+E]"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar. [Ctrl+G]"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar.[Ctrl+C]"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "vista"
            Object.ToolTipText     =   "Vista código o html"
            Style           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "genweb"
            Object.ToolTipText     =   "Generar Web. [Ctrl+W]"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "genintra"
            Object.ToolTipText     =   "Generar Intranet. [Ctrl+I]"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   400
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "preview"
            Object.ToolTipText     =   "Preview"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "edithtml"
            Object.ToolTipText     =   "Editor HTML.[Ctrl+L]"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "help"
            Object.ToolTipText     =   "Ayuda.[Ctrl+H]"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIconos 
      Left            =   5040
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatosWeb.frx":08CA
            Key             =   "salir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatosWeb.frx":0BE6
            Key             =   "nuevo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatosWeb.frx":0CFA
            Key             =   "codigo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatosWeb.frx":114E
            Key             =   "web"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatosWeb.frx":1A2A
            Key             =   "modificar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatosWeb.frx":1B86
            Key             =   "eliminar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatosWeb.frx":1C9A
            Key             =   "grabar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatosWeb.frx":1DAE
            Key             =   "cancelar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatosWeb.frx":1EC2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatosWeb.frx":21DE
            Key             =   "edithtml"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatosWeb.frx":24FA
            Key             =   "glosario"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatosWeb.frx":281A
            Key             =   "help"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatosWeb.frx":2B36
            Key             =   "medidas"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatosWeb.frx":2F8A
            Key             =   "genweb"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatosWeb.frx":32A6
            Key             =   "genintra"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatosWeb.frx":35C2
            Key             =   "negrita"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatosWeb.frx":36D6
            Key             =   "subrayar"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatosWeb.frx":37EA
            Key             =   "viñeta"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatosWeb.frx":38FE
            Key             =   "enter"
         EndProperty
      EndProperty
   End
   Begin VB.Label lUbicacionFoto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   900
      TabIndex        =   15
      Top             =   960
      Width           =   465
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Plantilla Intranet:"
      Height          =   195
      Left            =   4980
      TabIndex        =   8
      Top             =   1260
      Width           =   1395
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Plantilla Web:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1260
      Width           =   1155
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Prioridad:"
      Height          =   195
      Left            =   4980
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Foto:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Artículo:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   735
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuOpModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuOpEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuOpLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpGrabar 
         Caption         =   "Grabar"
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuOpCancelar 
         Caption         =   "&Cancelar"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuOpLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpGenera 
         Caption         =   "Generar Archivo &Html"
         Begin VB.Menu MnuOpGenWeb 
            Caption         =   "&Web"
            Shortcut        =   ^W
         End
         Begin VB.Menu MnuOpGenIntra 
            Caption         =   "&Intranet"
            Shortcut        =   ^I
         End
      End
      Begin VB.Menu MnuOpLinea3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "Salir"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuVer 
      Caption         =   "&Ver"
      Begin VB.Menu MnuOpPreview 
         Caption         =   "Preview"
      End
      Begin VB.Menu MnuOpVista 
         Caption         =   "Ver HTML"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu MnuInsertar 
      Caption         =   "&Insertar"
      Begin VB.Menu MnuInsHTML 
         Caption         =   "&Código HTML"
         Shortcut        =   ^L
      End
      Begin VB.Menu MnuInsGlosario 
         Caption         =   "&Glosario"
         Shortcut        =   ^O
      End
      Begin VB.Menu MnuInsLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuInsMedidas 
         Caption         =   "&Medidas"
      End
   End
   Begin VB.Menu MnuAyuda 
      Caption         =   "?"
      Begin VB.Menu MnuAyuHelp 
         Caption         =   "Ayuda"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmDatosWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sCaracteristica As String, sRecalcar As String
Private iFormatoWeb As Integer, iFormatoIntra As Integer

Private Sub Form_Load()
    IconosToolbar
    MiBotones False, False, False
    MenuTexto False
    lUbicacionFoto.Caption = Trim(pathFotos)
    ObtengoSeteoForm Me, 500, 500
    picTexto.ZOrder 0: tTexto.ZOrder 0
    LimpioObjetos
    If idArticulo > 0 Then BuscoArticuloPorCodigo 0, idArticulo
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    With tsTexto
        .Left = 20: tsTexto.Width = Me.ScaleWidth - 40
        .Height = Me.ScaleHeight - tsTexto.Top
    End With
    With picTexto
        .Top = tsTexto.ClientTop
        .Left = tsTexto.ClientLeft + 10: .Width = tsTexto.ClientWidth - 20
        .Height = tsTexto.ClientHeight - 20
    End With
    With picFoto
        .Top = tsTexto.ClientTop
        .Left = tsTexto.ClientLeft + 10: .Width = tsTexto.ClientWidth - 20
        .Height = tsTexto.ClientHeight - 20
    End With
    With imgFoto
        .Left = 10
        .Top = 10
    End With
    With vsFoto
        .Top = 0
        .Height = picFoto.Height - (50 + hsFoto.Height)
        .Left = picFoto.Width - (.Width + 50)
    End With
    With picScroll
        .Left = 0
        .Width = picFoto.Width '- (picFoto.Width - vsFoto.Left)
        .Top = picFoto.Height - (.Height + 50)
    End With
    With hsFoto
        .Width = picFoto.Width - (picFoto.Width - vsFoto.Left)
    End With
    With tooTexto
        .Left = 5: .Top = 5: .Width = picTexto.Width - 50
    End With
    With tTexto
        .Left = 10: .Top = tooMenu.Height + 5: .Width = picTexto.Width - 80
        .Height = picTexto.Height - (tooMenu.Height + 15)
    End With
    With tHtml
        .Top = tTexto.Top
        .Left = tTexto.Left
        .Width = tTexto.Width
        .Height = tTexto.Height
    End With
    With tFoto
        .Left = lUbicacionFoto.Left + lUbicacionFoto.Width + 50
        .Width = 2670
        If .Width + .Left < tPrioridad.Left + tPrioridad.Width Then
            .Width = (tPrioridad.Left + tPrioridad.Width) - .Left
        End If
    End With
    ManejoScroll
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    GuardoSeteoForm Me
    CierroConexion
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    End
End Sub

Private Sub hsFoto_Change()
    imgFoto.Left = -hsFoto.Value
End Sub

Private Sub Label1_Click()
On Error Resume Next
    With tArticulo
        If Not .Enabled Then Exit Sub
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label2_Click()
On Error Resume Next
    With tFoto
        If Not .Enabled Then Exit Sub
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label3_Click()
On Error Resume Next
    With tPrioridad
        If Not .Enabled Then Exit Sub
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label4_Click()
On Error Resume Next
    With tWEB
        If Not .Enabled Then Exit Sub
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub Label5_Click()
On Error Resume Next
    With tIntranet
        If Not .Enabled Then Exit Sub
        .SelStart = 0: .SelLength = Len(.Text): .SetFocus
    End With
End Sub

Private Sub MnuAyuHelp_Click()
    AccionAyuda
End Sub

Private Sub MnuInsGlosario_Click()
    AccionGlosario
End Sub

Private Sub MnuInsHTML_Click()
    AccionHTML
End Sub

Private Sub MnuInsMedidas_Click()
    AccionMedidas
End Sub

Private Sub MnuOpEliminar_Click()
    AccionEliminar
End Sub

Private Sub MnuOpGenIntra_Click()
    AccionGenerarIntra
End Sub

Private Sub MnuOpGenWeb_Click()
    AccionGenerarWeb
End Sub

Private Sub MnuOpGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuOpModificar_Click()
    AccionModificar
End Sub

Private Sub MnuOpPreview_Click()
    AccionPreview
End Sub

Private Sub MnuOpVista_Click()
    MnuOpVista.Checked = Not MnuOpVista.Checked
    If tooMenu.Buttons("vista").Value = tbrPressed Then
        tooMenu.Buttons("vista").Value = tbrUnpressed
    Else
        tooMenu.Buttons("vista").Value = tbrPressed
    End If
    AccionVista
End Sub

Private Sub MnuSalir_Click()
    Unload Me
End Sub

Private Sub IconosToolbar()
    With tooMenu
        .ImageList = imgIconos
        .Buttons("salir").Image = imgIconos.ListImages("salir").Index
        .Buttons("modificar").Image = imgIconos.ListImages("modificar").Index
        .Buttons("eliminar").Image = imgIconos.ListImages("eliminar").Index
        .Buttons("grabar").Image = imgIconos.ListImages("grabar").Index
        .Buttons("cancelar").Image = imgIconos.ListImages("cancelar").Index
        If .Buttons("vista").Value = tbrUnpressed Then
            .Buttons("vista").Image = imgIconos.ListImages("codigo").Index
        Else
            .Buttons("vista").Image = imgIconos.ListImages("web").Index
        End If
        .Buttons("preview").Image = imgIconos.ListImages("preview").Index
        .Buttons("help").Image = imgIconos.ListImages("help").Index
        .Buttons("edithtml").Image = imgIconos.ListImages("edithtml").Index
        .Buttons("genweb").Image = imgIconos.ListImages("genweb").Index
        .Buttons("genintra").Image = imgIconos.ListImages("genintra").Index
    End With
    With tooTexto
        .ImageList = imgIconos
        .Buttons("enter").Image = imgIconos.ListImages("enter").Index
        .Buttons("negrita").Image = imgIconos.ListImages("negrita").Index
        .Buttons("viñeta").Image = imgIconos.ListImages("viñeta").Index
        .Buttons("subrayar").Image = imgIconos.ListImages("subrayar").Index
        .Buttons("glosario").Image = imgIconos.ListImages("glosario").Index
        .Buttons("medidas").Image = imgIconos.ListImages("medidas").Index
    End With
End Sub

Private Sub MiBotones(ByVal bModificar As Boolean, ByVal bEliminar As Boolean, ByVal bGrabar As Boolean)
    
    With tooMenu
        .Buttons("modificar").Enabled = bModificar
        .Buttons("eliminar").Enabled = bEliminar
        .Buttons("grabar").Enabled = bGrabar
        .Buttons("cancelar").Enabled = bGrabar
        .Buttons("preview").Enabled = bEliminar
        .Buttons("edithtml").Enabled = False
        If Val(tArticulo.Tag) > 0 Then
            .Buttons("genweb").Enabled = True
            .Buttons("genintra").Enabled = True
        Else
            .Buttons("genweb").Enabled = False
            .Buttons("genintra").Enabled = False
        End If
    End With
    
    MenuTexto False
    
    MnuOpModificar.Enabled = bModificar
    MnuOpEliminar.Enabled = bEliminar
    MnuOpGrabar.Enabled = bGrabar
    MnuOpCancelar.Enabled = bGrabar
    MnuOpPreview.Enabled = bEliminar
    
    If Val(tArticulo.Tag) > 0 Then
        MnuOpGenIntra.Enabled = True
        MnuOpGenWeb.Enabled = True
    Else
        MnuOpGenIntra.Enabled = False
        MnuOpGenWeb.Enabled = False
    End If
    
    MnuInsHTML.Enabled = False
    
    EstadoObjetos bGrabar
    
    
End Sub

Private Sub EstadoObjetos(ByVal bHabilitar As Boolean)

    tArticulo.Enabled = Not bHabilitar
    tPrioridad.Enabled = bHabilitar
    tFoto.Enabled = bHabilitar
    tWEB.Enabled = bHabilitar
    tIntranet.Enabled = bHabilitar
    tTexto.Locked = Not bHabilitar
    
    If bHabilitar Then
        tArticulo.BackColor = vbButtonFace
        tPrioridad.BackColor = &HC0FFFF
        tFoto.BackColor = vbWindowBackground
        tWEB.BackColor = vbWindowBackground
        tIntranet.BackColor = vbWindowBackground
    Else
        tArticulo.BackColor = vbWindowBackground
        tPrioridad.BackColor = vbButtonFace
        tFoto.BackColor = vbButtonFace
        tWEB.BackColor = vbButtonFace
        tIntranet.BackColor = vbButtonFace
    End If
                    
End Sub

Private Sub tArticulo_Change()
    If Val(tArticulo.Tag) > 0 Then
        MiBotones False, False, False
        tHtml.DocumentHTML = ""
        LimpioObjetos
        tArticulo.Tag = ""
        tArticulo.SetFocus
    End If
End Sub

Private Sub tArticulo_GotFocus()
On Error Resume Next
    With tArticulo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tArticulo_KeyPress(KeyAscii As Integer)
On Error GoTo errAP
    Screen.MousePointer = 11
    If KeyAscii = vbKeyReturn Then
        If Trim(tArticulo.Text) <> "" Then
            If IsNumeric(tArticulo.Text) Then
                BuscoArticuloPorCodigo tArticulo.Text
            Else
                BuscoArticuloPorNombre tArticulo.Text
            End If
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
errAP:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub BuscoArticuloPorCodigo(CodArticulo As Long, Optional IDArt As Long = 0)
'Atención el mapeo de error lo hago antes de entrar al procedimiento
    Screen.MousePointer = 11
    If IDArt = 0 Then
        Cons = "Select * From Articulo Where ArtCodigo = " & CodArticulo
    Else
        Cons = "Select * From Articulo Where Artid = " & IDArt
    End If
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        tArticulo.Text = "": tArticulo.Tag = "0"
        MsgBox "No existe un artículo que posea ese código.", vbExclamation, "ATENCIÓN"
    Else
        tArticulo.Text = Format(RsAux!ArtCodigo, "#,000,000") & " " & Trim(RsAux!ArtNombre)
        tArticulo.Tag = RsAux!ArtID
        RsAux.Close
        CargoDatosWeb
    End If
    Screen.MousePointer = 0
End Sub

Private Sub BuscoArticuloPorNombre(NomArticulo As String)
'Atención el mapeo de error lo hago antes de entrar al procedimiento
Dim Resultado As Long

    Screen.MousePointer = 11
    Resultado = 0
    Cons = "Select ArtId, Código = ArtCodigo, Nombre = ArtNombre from Articulo" _
        & " Where ArtNombre LIKE '" & clsGeneral.Replace(NomArticulo, " ", "%") & "%'" _
        & " Order By ArtNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        tArticulo.Tag = "0"
        Screen.MousePointer = 0
        MsgBox "No existe un nombre de artículo que concuerde con los datos ingresados.", vbExclamation, "ATENCIÓN"
    Else
        RsAux.MoveNext
        If RsAux.EOF Then
            RsAux.MoveFirst
            Resultado = RsAux!Código
            RsAux.Close
        Else
            RsAux.Close
            Screen.MousePointer = 11
            Dim objAyuda As New clsListadeAyuda
            If objAyuda.ActivarAyuda(cBase, Cons, 5000, 1, "Lista de Artículos") > 0 Then
                Resultado = objAyuda.RetornoDatoSeleccionado(1)
            End If
            Set objAyuda = Nothing
        End If
    End If
    If Resultado > 0 Then BuscoArticuloPorCodigo Resultado
    Screen.MousePointer = 0
End Sub

Private Sub CargoDatosWeb()
On Error GoTo errCDW
    Screen.MousePointer = 11
    LimpioObjetos
    tHtml.DocumentHTML = ""
    Cons = "Select ArticuloWebPage.*, PlaWeb.PlaNombre As WebNombre, PlaWeb.PlaFormato As WebFormato, PlaIntra.PlaNombre As IntrNombre, PlaIntra.PlaFormato As IntrFormato" _
        & " From ArticuloWebPage " _
            & " Left Outer Join Plantilla PlaWeb ON PlaWeb.PlaCodigo = AWPPlantillaWeb" _
            & " Left Outer Join Plantilla PlaIntra ON  PlaINtra.PlaCodigo = AWPPlantillaIntra" _
        & " Where AWPArticulo = " & Val(tArticulo.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        MiBotones True, True, False
        If Not IsNull(RsAux!AWPFoto) Then tFoto.Text = Trim(RsAux!AWPFoto)
        tPrioridad.Text = Trim(RsAux!AWPPrioridad)
        If Not IsNull(RsAux!AWPPlantillaWeb) Then
            tWEB.Text = Trim(RsAux!WebNombre)
            tWEB.Tag = RsAux!AWPPlantillaWeb
            iFormatoWeb = RsAux!WebFormato
        End If
        If Not IsNull(RsAux!AWPPlantillaIntra) Then
            iFormatoIntra = RsAux!IntrFormato
            tIntranet.Text = Trim(RsAux!IntrNombre)
            tIntranet.Tag = RsAux!AWPPlantillaIntra
        End If
        sCaracteristica = LongTextDeRSAUX("AWPTexto")
        sRecalcar = LongTextDeRSAUX("AWPTexto2")
        tooMenu.Buttons("vista").Value = tbrUnpressed
        Select Case tsTexto.SelectedItem.Key
            Case "datos"
                tTexto.ZOrder 0
                tTexto.Text = sRecalcar
            Case Else
                tTexto.ZOrder 0
                tTexto.Text = sCaracteristica
        End Select
        If Trim(tFoto.Text) <> "" Then 'Cargo la foto
            CargoFoto False
        End If
    Else
        MiBotones True, False, False
    End If
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub
errCDW:
    clsGeneral.OcurrioError "Ocurrió el siguiente error al cargar los datos del artículo.", Err.Description, "Error (cargodatosweb)"
End Sub

Private Function LongTextDeRSAUX(ByVal sColumna As String) As String
On Error GoTo errLT
    LongTextDeRSAUX = ""
    LongTextDeRSAUX = RsAux(sColumna)
errFin:
    Exit Function
errLT:
    Resume errFin
End Function

Private Sub tFoto_GotFocus()
    With tFoto
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tFoto_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If Trim(tFoto.Text) <> "" Then CargoFoto True
        tWEB.SetFocus
    End If
End Sub


Private Sub tIntranet_Change()
    tIntranet.Tag = ""
End Sub

Private Sub tIntranet_GotFocus()
On Error Resume Next
    With tIntranet
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tIntranet_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim sNombre As String, lCod As Long
    If KeyAscii = vbKeyReturn Then
        If Trim(tIntranet.Text) <> "" And Val(tIntranet.Tag) = 0 Then
            'Busco las plantillas del tipo 7.
            sNombre = tIntranet.Text
            clsGeneral.Replace sNombre, " ", "%"
            lCod = BuscoPlantilla(sNombre, iFormatoIntra)
            If Val(lCod) > 0 Then
                tIntranet.Text = sNombre
                tIntranet.Tag = lCod
            End If
        End If
        tsTexto.SetFocus
    End If
End Sub

Private Sub tooMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "salir": Unload Me
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        Case "glosario": AccionGlosario
        Case "preview": AccionPreview
        Case "vista": MnuOpVista.Checked = Not MnuOpVista.Checked: AccionVista
        Case "edithtml": AccionHTML
        Case "medidas": AccionMedidas
        Case "help": AccionAyuda
        Case "genweb": AccionGenerarWeb
        Case "genintra": AccionGenerarIntra
    End Select
End Sub

Private Sub tooTexto_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "glosario": AccionGlosario
        Case "medidas": AccionMedidas
        Case "negrita": AplicoNegrita
        Case "viñeta": AplicoViñeta
        Case "subrayar": AplicoSubrayar
        Case "enter": AplicoEnter
    End Select
End Sub

Private Sub tPrioridad_GotFocus()
On Error Resume Next
    With tPrioridad
        .SelStart = 0: .SelLength = Len(.Text)
        .SetFocus
    End With
End Sub

Private Sub tPrioridad_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tPrioridad.Text) Then
            tFoto.SetFocus
        Else
            MsgBox "Debe ingresar la prioridad.", vbExclamation, "ATENCIÓN"
        End If
    End If
End Sub

Private Sub tsTexto_Click()
    Select Case tsTexto.SelectedItem.Key
        Case "foto": picFoto.ZOrder 0
        Case "caracteristicas"
            tTexto.Text = sCaracteristica
            If sCaracteristica <> "" Then
                tHtml.DocumentHTML = tTexto.Text
                Do While tHtml.Busy
                    DoEvents
                Loop
            End If
            If tooMenu.Buttons("vista").Value = tbrUnpressed Then tTexto.ZOrder 0 Else tHtml.ZOrder 0
            picTexto.ZOrder 0
        Case "datos"
            
            tTexto.Text = sRecalcar
            If sRecalcar <> "" Then
                tHtml.DocumentHTML = tTexto.Text
                Do While tHtml.Busy
                    DoEvents
                Loop
            End If
            If tooMenu.Buttons("vista").Value = tbrUnpressed Then tTexto.ZOrder 0 Else tHtml.ZOrder 0
            picTexto.ZOrder 0
    End Select
End Sub

Private Function BuscoPlantilla(ByRef sNombre As String, ByRef iFormato As Integer) As Long
On Error GoTo errBP
    BuscoPlantilla = 0: iFormato = 0
    Screen.MousePointer = 11
    Cons = "Select PlaCodigo as 'Código', PlaNombre as 'Nombre', PlaFormato as 'Formato' " _
        & " From Plantilla Where PlaNombre Like '" & sNombre & "%' And PlaTipo = 8"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        MsgBox "No existe una plantilla con ese nombre.", vbExclamation, "ATENCIÓN"
    Else
        RsAux.MoveNext
        If RsAux.EOF Then
            RsAux.MoveFirst
            sNombre = Trim(RsAux!Nombre)
            iFormato = RsAux!Formato
            BuscoPlantilla = RsAux!Código
            RsAux.Close
        Else
            RsAux.Close
            Dim objAyuda As New clsListadeAyuda
            If objAyuda.ActivarAyuda(cBase, Cons, 5000, 0, "Lista de plantillas") > 0 Then
                Cons = "Select * From Plantilla Where PlaCodigo = " & objAyuda.RetornoDatoSeleccionado(0)
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsAux.EOF Then
                    sNombre = Trim(RsAux!PlaNombre)
                    iFormato = RsAux!PlaFormato
                    BuscoPlantilla = RsAux!PlaCodigo
                End If
                RsAux.Close
            End If
            Set objAyuda = Nothing
        End If
    End If
    Screen.MousePointer = 0
    Exit Function
errBP:
    clsGeneral.OcurrioError "Ocurrió un error al buscar la plantilla.", Err.Description, "Error (buscoplantilla)"
    Screen.MousePointer = 0
End Function
Private Sub AccionModificar()
    tHtml.BrowseMode = False
    If tooMenu.Buttons("vista").Value = tbrPressed Then
        tooMenu.Buttons("vista").Value = tbrUnpressed
        MnuOpVista.Checked = Not MnuOpVista.Checked
    End If
    AccionVista
    MiBotones False, False, True
'    MenuTexto True
    Me.Refresh
    tPrioridad.SetFocus
End Sub
Private Sub AccionEliminar()
On Error GoTo errAE
    
    If MsgBox("¿Confirma eliminar los datos web para el artículo?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
        
        FechaDelServidor
        
        Cons = "Select * From Articulo Where ArtID = " & Val(tArticulo.Tag)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.Edit
        RsAux!ArtModificado = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
        RsAux.Update
        RsAux.Close
    
        Cons = "Select * From ArticuloWebPage Where AWPArticulo = " & Val(tArticulo.Tag)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.Delete
        RsAux.Close
        LimpioObjetos
        tHtml.DocumentHTML = ""
        tArticulo.Text = ""
        MiBotones False, False, False
    End If
    Exit Sub
    
errAE:
    clsGeneral.OcurrioError "Ocurrió un error al intentar eliminar el registro.", Err.Description
End Sub
Private Sub AccionCancelar()
    MenuTexto False
    CargoDatosWeb
End Sub
Private Sub AccionGrabar()
    If Not ValidoDatos Then Exit Sub
    If MsgBox("¿Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
        GraboDatos
    End If
End Sub
Private Sub AccionGlosario()
On Error GoTo errAG
Dim sAux As String, sTexto As String
Dim objLista As New clsListadeAyuda
    
    Cons = "Select GloID, GloAlto, GloAncho, GloScroll, GloNombre As 'Nombre' From Glosario Order by GloNombre"
    If objLista.ActivarAyuda(cBase, Cons, 3500, 4, "Lista de Glosarios ") > 0 Then
        
        sTexto = clsGeneral.Replace(tTexto.SelText, vbCrLf, "")
        sAux = "<a href=@glosario.asp?Id=" & objLista.RetornoDatoSeleccionado(0) & _
                "@  onclick=@NewWindow(this.href, 'name','" & objLista.RetornoDatoSeleccionado(2) & _
                "','" & objLista.RetornoDatoSeleccionado(1) & "','" & IIf(objLista.RetornoDatoSeleccionado(3), "Yes", "No") & "');return false;@>"
                
        If Trim(tTexto.SelText) = "" Then sAux = sAux & Trim$(objLista.RetornoDatoSeleccionado(4))
                
        sAux = clsGeneral.Replace(sAux, "@", Chr(34))       'Truchada.
        tTexto.SelText = sAux & sTexto
    End If
    Exit Sub
    
errAG:
    clsGeneral.OcurrioError "Ocurrió un error al intentar acceder a la lista de glosarios.", Err.Description
End Sub
Private Sub AccionPreview()
On Error GoTo errAP
    
    Dim objPreview As New clsPlantillaI
    
    Select Case tsTexto.SelectedItem.Key
        Case "foto"
            MsgBox "Campo sin preview, seleccione 'Características' ó 'Datos a Remarcar'", vbInformation, "ATENCIÓN"
        Case "datos"
            If Val(tIntranet.Tag) > 0 Then
                objPreview.ProcesoPlantillaInteractiva cBase, Val(tIntranet.Tag), iFormatoIntra, "", "", Val(tArticulo.Tag), True
            Else
                MsgBox "Debe seleccionar una plantilla para la INTRANET.", vbInformation, "ATENCIÓN"
            End If
            
        Case Else
            If Val(tWEB.Tag) > 0 Then
                objPreview.ProcesoPlantillaInteractiva cBase, Val(tWEB.Tag), iFormatoWeb, "", "", Val(tArticulo.Tag), True
            Else
                MsgBox "Debe seleccionar una plantilla para la WEB.", vbInformation, "ATENCIÓN"
            End If
    End Select
    
    Set objPreview = Nothing
    Exit Sub
    
errAP:
    clsGeneral.OcurrioError "Ocurrió un error al instanciar el preview.", Err.Description, "Preview"
End Sub
Private Sub tTexto_Change()
On Error Resume Next

    Select Case tsTexto.SelectedItem.Key
        Case "caracteristicas"
            sCaracteristica = tTexto.Text
        Case "datos"
            sRecalcar = tTexto.Text
    End Select
    
End Sub

Private Sub tTexto_GotFocus()
    If Not tTexto.Locked Then MenuTexto True
End Sub

Private Sub tTexto_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 45 Then            'shift + insert
        If Not tTexto.Locked Then
            If Trim(Clipboard.GetText) <> "" Then AplicoFormatoHTML
        End If
    End If
    
End Sub

Private Sub tTexto_LostFocus()
    MenuTexto False
End Sub

Private Sub tWEB_Change()
    tWEB.Tag = ""
End Sub

Private Sub tWEB_GotFocus()
    With tWEB
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tWEB_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim sNombre As String, lCod As Long
    If KeyAscii = vbKeyReturn Then
        If Trim(tWEB.Text) <> "" And Val(tWEB.Tag) = 0 Then
            'Busco las plantillas del tipo 7.
            sNombre = tWEB.Text
            clsGeneral.Replace sNombre, " ", "%"
            lCod = BuscoPlantilla(sNombre, iFormatoWeb)
            If lCod > 0 Then
                tWEB.Text = sNombre
                tWEB.Tag = lCod
            End If
        End If
        tIntranet.SetFocus
    End If
End Sub

Private Sub GraboDatos()
On Error GoTo errGD
    Screen.MousePointer = 11
    
    FechaDelServidor
    Cons = "Select * From Articulo Where ArtID = " & Val(tArticulo.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.Edit
    RsAux!ArtModificado = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    RsAux.Update
    RsAux.Close
    
    Cons = "Select * From ArticuloWebPage Where AWPArticulo = " & Val(tArticulo.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        RsAux.Edit
    Else
        RsAux.AddNew
    End If
    RsAux!AWPArticulo = Val(tArticulo.Tag)
    RsAux!AWPPrioridad = Val(tPrioridad.Text)
    If Trim(tFoto.Text) <> "" Then RsAux!AWPFoto = Trim(tFoto.Text) Else RsAux!AWPFoto = Null
    If Val(tWEB.Tag) > 0 Then RsAux!AWPPlantillaWeb = Val(tWEB.Tag) Else RsAux!AWPPlantillaWeb = Null
    If Val(tIntranet.Tag) > 0 Then RsAux!AWPPlantillaIntra = Val(tIntranet.Tag) Else RsAux!AWPPlantillaIntra = Null
    If Trim(sCaracteristica) = "" Then RsAux!AWPTexto = Null Else RsAux!AWPTexto = sCaracteristica
    If Trim(sRecalcar) = "" Then RsAux!AWPTexto2 = Null Else RsAux!AWPTexto2 = sRecalcar
    RsAux.Update
    RsAux.Close
    AccionCancelar
    Screen.MousePointer = 0
    Exit Sub
errGD:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió el siguiente error al almacenar la información.", Err.Description, "Grabar"
End Sub

Private Function ValidoDatos() As Boolean
    ValidoDatos = False
    If Not IsNumeric(tPrioridad.Text) Then
        MsgBox "Debe ingresar un número de prioridad.", vbExclamation, "ATENCIÓN"
        tPrioridad.SetFocus: Exit Function
    End If
    If Val(tWEB.Tag) = 0 Then
        MsgBox "No ingreso una plantilla para la web, si los datos se presentarán o actualizaran por el CENTINELA debe necesariamente ingresar una.", vbInformation, "ATENCIÓN"
    End If
    If Val(tIntranet.Tag) = 0 Then
        MsgBox "No ingreso una plantilla para la intranet, si los datos se presentarán o actualizaran por el CENTINELA debe necesariamente ingresar una.", vbInformation, "ATENCIÓN"
    End If
    
    If Trim(tFoto.Text) <> "" Then
        On Error Resume Next
        If Dir(pathFotos & Trim(tFoto.Text), vbArchive) = "" Then
            MsgBox "No se encontró la foto, verifique.", vbExclamation, "ATENCIÓN"
        End If
    End If
    
    ValidoDatos = True
End Function
Private Sub CargoFoto(ByVal bMsgError As Boolean)
On Error GoTo errCF
    imgFoto.Picture = LoadPicture(pathFotos & tFoto.Text)
    ManejoScroll
    Exit Sub
errCF:
    If bMsgError Then clsGeneral.OcurrioError "Ocurrió el siguiente error al cargar la foto.", Err.Description, "Cargo Foto"
End Sub
Private Sub AccionHTML()
On Error GoTo errED
    Dim objEditor As New clsEditHTML
    Dim sAux As String
    sAux = objEditor.EditorHTML(tTexto.Text)
    Set objEditor = Nothing
    If sAux <> "" Then
        tTexto.Text = ""        'go to change.
        tTexto.Text = sAux
    End If
    Exit Sub
errED:
    clsGeneral.OcurrioError "Ocurrió el siguietne error al activar el editor.", Err.Description, "Editor HTML"
End Sub
Private Sub LimpioObjetos()
    
    tPrioridad.Text = ""
    tFoto.Text = ""
    tTexto.Text = ""
    tWEB.Text = ""
    tIntranet.Text = ""
    tooMenu.Buttons("vista").Value = tbrUnpressed
    tooMenu.Buttons("vista").Image = imgIconos.ListImages("codigo").Index
    imgFoto.Picture = LoadPicture()
    tHtml.Visible = False
    
    'Inicializo Globales
    sCaracteristica = "": sRecalcar = ""
    iFormatoWeb = 0: iFormatoIntra = 0
    
    imgFoto.Left = 0: imgFoto.Top = 0
    vsFoto.Enabled = False: hsFoto.Enabled = False
    vsFoto.Value = 0: vsFoto.Max = 0
    
End Sub

Private Sub AccionVista()
On Error GoTo errAV
    
    Select Case tooMenu.Buttons("vista").Value
        Case tbrUnpressed
            tHtml.Visible = False
            tooMenu.Buttons("vista").Image = imgIconos.ListImages("codigo").Index
            Select Case tsTexto.SelectedItem.Key
                Case "foto": Exit Sub
                Case "caracteristicas"
                    tTexto.ZOrder 0
                    tTexto.Text = sCaracteristica
                Case "datos"
                    tTexto.ZOrder 0
                    tTexto.Text = sRecalcar
            End Select
            If Not tTexto.Locked Then tTexto.SetFocus
        Case Else
            tooMenu.Buttons("vista").Image = imgIconos.ListImages("web").Index
            Select Case tsTexto.SelectedItem.Key
                Case "foto": Exit Sub
                Case "caracteristicas"
                    If iFormatoWeb = 1 Then
                        MsgBox "La plantilla seleccionada para la web es plana.", vbInformation, "ATENCIÓN"
                        tooMenu.Buttons("vista").Value = tbrUnpressed
                        Exit Sub
                    End If
                    With tHtml
                        .Visible = True
                        .ZOrder 0
                        .DocumentHTML = sCaracteristica
                    End With
                
                Case "datos"
                    If iFormatoIntra = 1 Then
                        MsgBox "La plantilla seleccionada para la intranet es plana.", vbInformation, "ATENCIÓN"
                        tooMenu.Buttons("vista").Value = tbrUnpressed
                        Exit Sub
                    End If
                    With tHtml
                        .Visible = True
                        .ZOrder 0
                        .DocumentHTML = sRecalcar
                    End With
                    tsTexto.SetFocus
            End Select
    End Select
    Exit Sub
errAV:
    MsgBox "Error al activar vista previa." & vbCrLf & vbCrLf & "Error: " & Err.Description, vbCritical, "Atención"
End Sub

Private Sub AccionAyuda()
On Error GoTo errHelp

    Screen.MousePointer = 11
    Dim aFile As String
    Cons = "Select * from Aplicacion Where AplNombre = '" & App.Title & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux!AplHelp) Then aFile = Trim(RsAux!AplHelp)
    RsAux.Close
    
    If aFile <> "" Then EjecutarApp aFile
    
    Screen.MousePointer = 0
    Exit Sub
    
errHelp:
    clsGeneral.OcurrioError "Error al activar el archivo de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionMedidas()
On Error GoTo errAM
Dim sAux As String
    'Busco el texto, si esta ingresado doy advertencia
    If LCase(tTexto.Text) Like "*medidas*#x#*" Then
        If MsgBox("Ya tiene las medidas.  ¿Las agrega igualmente?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
            Exit Sub
        End If
    End If
    Screen.MousePointer = 11
    sAux = ""
    
    Cons = "Select * From ArticuloFacturacion Where AFaArticulo = " & Val(tArticulo.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!AFaAlto) Then sAux = CStr(RsAux!AFaAlto) & "x"
        If Not IsNull(RsAux!AFaFrente) Then sAux = sAux & CStr(RsAux!AFaFrente) & "x"
        If Not IsNull(RsAux!AFaProfundidad) Then sAux = sAux & CStr(RsAux!AFaProfundidad) & "x"
    End If
    RsAux.Close
    
    If sAux <> "" Then
        'Cambio el último x por cm.
        sAux = Mid(sAux, 1, Len(sAux) - 1) & "cm"
                
        If Mid(tTexto.Text, 1, tTexto.SelStart) <> "" Then
            If InStr(1, tTexto.Text, "&#149;") Then
                sAux = "<BR> &#149; <U>Medidas</U>: " & sAux
            Else
                sAux = "<BR> <U>Medidas</U>: " & sAux
            End If
        Else
            sAux = "<U>Medidas</U>: " & sAux
        End If
        tTexto.SelText = sAux
    Else
        MsgBox "No hay datos de medidas ingresados para este artículo.", vbInformation, "ATENCIÓN"
    End If
    Screen.MousePointer = 0
    Exit Sub
errAM:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió el siguiente error al agregar las medidas.", Err.Description, "Agregar Medidas"
End Sub

Private Sub AccionGenerarWeb()
    If Val(tWEB.Tag) = 0 Then
        MsgBox "Debe seleccionar una plantilla para la web.", vbInformation, "ATENCIÓN"
        Exit Sub
    End If
    AccionGenerarArchivo Val(tWEB.Tag), 1
End Sub

Private Sub AccionGenerarIntra()
    If Val(tIntranet.Tag) = 0 Then
        MsgBox "Debe seleccionar una plantilla para la intranet.", vbInformation, "ATENCIÓN"
        Exit Sub
    End If
    AccionGenerarArchivo Val(tIntranet.Tag), 2
End Sub

Private Sub AccionGenerarArchivo(ByVal idPlantilla As Long, ByVal iTipo As Integer)
Dim objPlantilla As New clsPlantillaI
Dim iFormato As Integer, sTexto As String, sAsunto As String
    
    On Error GoTo errAGA
    If MsgBox("¿Confirma generar el archivo html?", vbQuestion + vbYesNo, "Generar") = vbNo Then Exit Sub
    sTexto = ""
    sAsunto = ""
    If objPlantilla.ProcesoPlantillaInteractiva(cBase, idPlantilla, iFormato, sTexto, sAsunto, CStr(Val(tArticulo.Tag)), False) Then
        If iFormato = 1 Then
            If MsgBox("El formato de la plantilla generada no es html." & vbCrLf & "¿Confirma generar la página html?", vbQuestion + vbYesNo, "Posible Error") = vbNo Then
                GoTo evSalir
            End If
        End If
        If sAsunto = "" Then
            MsgBox "No se obtuvo el nombre del archivo a almacenar, verifique en la plantilla si tiene asuntoRef.", vbCritical, "ATENCIÓN"
            GoTo evSalir
        Else
            If iTipo = 1 Then
                sAsunto = Replace(pathWeb, "[id]", Val(tArticulo.Tag), , , vbTextCompare)
            Else
                sAsunto = Replace(pathIntra, "[id]", Val(tArticulo.Tag), , , vbTextCompare)
            End If
            Open sAsunto For Output As #1
            Print #1, sTexto
            Close #1
        End If
    Else
        MsgBox "Ocurrió un error al procesar la plantilla, verifique con el preview.", vbExclamation, "ATENCIÓN"
        GoTo evSalir
    End If

evSalir:
    Set objPlantilla = Nothing
    Exit Sub
errAGA:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al intentar generar la página.", Err.Description, "Generar Archivo"
End Sub

Private Sub ManejoScroll()
    If imgFoto.Height > picScroll.Top Then
        vsFoto.Enabled = True
        vsFoto.Max = imgFoto.Height - picScroll.Top
    Else
        vsFoto.Enabled = False
        vsFoto.Max = 0
    End If
    If imgFoto.Width > vsFoto.Left Then
        hsFoto.Enabled = True
        hsFoto.Max = imgFoto.Width - vsFoto.Left
    Else
        hsFoto.Enabled = False
        hsFoto.Max = 0
    End If
End Sub

Private Sub vsFoto_Change()
    imgFoto.Top = -vsFoto.Value
End Sub

Private Sub MenuTexto(ByVal bHabilito As Boolean)
    
    With tooTexto
        .Buttons("glosario").Enabled = bHabilito
        .Buttons("medidas").Enabled = bHabilito
        .Buttons("viñeta").Enabled = bHabilito
        .Buttons("negrita").Enabled = bHabilito
        .Buttons("subrayar").Enabled = bHabilito
        .Buttons("enter").Enabled = bHabilito
    End With
        
    tooMenu.Buttons("edithtml").Enabled = bHabilito
    MnuInsHTML.Enabled = bHabilito
    MnuInsGlosario.Enabled = bHabilito
    MnuInsMedidas.Enabled = bHabilito
    
End Sub

Private Sub AplicoEnter()
    tTexto.SelText = tTexto.SelText & "<BR>"
End Sub

Private Sub AplicoViñeta()
    tTexto.SelText = "&#149;" & tTexto.SelText
End Sub

Private Sub AplicoNegrita()
    tTexto.SelText = "<B>" & tTexto.SelText & "</B>"
End Sub

Private Sub AplicoSubrayar()
    tTexto.SelText = "<U>" & tTexto.SelText & "</U>"
End Sub

Private Sub AplicoFormatoHTML()
    'Sustituyo los vbcrlf x <BR> & VBCRLF
    'Le agrego viñetas
    If InStr(1, tTexto.Text, "<BR>") = 0 Then
        If MsgBox("El texto que ingreso aparentemente es plano" & vbCrLf & "¿Desea asignarle las viñetas y los retorno de línea en formato HTML?", vbQuestion + vbYesNo, "Dar Formato") = vbYes Then
            tTexto.Text = "&#149" & tTexto.Text
            tTexto.Text = clsGeneral.Replace(tTexto.Text, vbCrLf, "<BR>" & vbCrLf & "&#149", , , vbTextCompare)
        End If
    End If
End Sub
