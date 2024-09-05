VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form frmParametros 
   Caption         =   "Parámetros Comercio"
   ClientHeight    =   7725
   ClientLeft      =   1935
   ClientTop       =   1230
   ClientWidth     =   10935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmParametros.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   10935
   Begin VB.CommandButton bSubrubroDifCostoImp 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   77
      Tag             =   "SubrubroDifCostoImp"
      Top             =   7080
      Width           =   375
   End
   Begin VB.CommandButton bSubrubroDifCambioG 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":0744
      Style           =   1  'Graphical
      TabIndex        =   74
      Tag             =   "SubrubroDifCambioG"
      Top             =   6720
      Width           =   375
   End
   Begin VB.CommandButton bSubrubroDifCambio 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":0A46
      Style           =   1  'Graphical
      TabIndex        =   71
      Tag             =   "SubrubroDifCambio"
      Top             =   6360
      Width           =   375
   End
   Begin VB.CommandButton bSubRubroVentas 
      Height          =   315
      Left            =   10440
      Picture         =   "frmParametros.frx":0D48
      Style           =   1  'Graphical
      TabIndex        =   68
      Tag             =   "SubRubroVentas"
      Top             =   6060
      Width           =   375
   End
   Begin VB.CommandButton bSubRubroIngresosVarios 
      Height          =   315
      Left            =   10440
      Picture         =   "frmParametros.frx":104A
      Style           =   1  'Graphical
      TabIndex        =   63
      Tag             =   "SubRubroIngresosVarios"
      Top             =   5340
      Width           =   375
   End
   Begin VB.CommandButton bSubRubroIva 
      Height          =   315
      Left            =   10440
      Picture         =   "frmParametros.frx":134C
      Style           =   1  'Graphical
      TabIndex        =   62
      Tag             =   "SubRubroIva"
      Top             =   5700
      Width           =   375
   End
   Begin VB.CommandButton bSubrubroVContado 
      Height          =   315
      Left            =   10440
      Picture         =   "frmParametros.frx":164E
      Style           =   1  'Graphical
      TabIndex        =   51
      Tag             =   "SubrubroVContado"
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton bSubrubroNDevolucion 
      Height          =   315
      Left            =   10440
      Picture         =   "frmParametros.frx":1950
      Style           =   1  'Graphical
      TabIndex        =   50
      Tag             =   "SubrubroNDevolucion"
      Top             =   4560
      Width           =   375
   End
   Begin VB.CommandButton bSubrubroCMorosidades 
      Height          =   315
      Left            =   10440
      Picture         =   "frmParametros.frx":1C52
      Style           =   1  'Graphical
      TabIndex        =   49
      Tag             =   "SubrubroCMorosidades"
      Top             =   4200
      Width           =   375
   End
   Begin VB.CommandButton bSubrubroCCuotas 
      Height          =   315
      Left            =   10440
      Picture         =   "frmParametros.frx":1F54
      Style           =   1  'Graphical
      TabIndex        =   48
      Tag             =   "SubrubroCCuotas"
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton bSubrubroNEspecial 
      Height          =   315
      Left            =   10440
      Picture         =   "frmParametros.frx":2256
      Style           =   1  'Graphical
      TabIndex        =   47
      Tag             =   "SubrubroNEspecial"
      Top             =   4920
      Width           =   375
   End
   Begin VB.CommandButton bMCSenias 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":2558
      Style           =   1  'Graphical
      TabIndex        =   44
      Tag             =   "MCSenias"
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton bSubrubroSeniasRecibidas 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":285A
      Style           =   1  'Graphical
      TabIndex        =   41
      Tag             =   "SubrubroSeniasRecibidas"
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton bMCTransferencias 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":2B5C
      Style           =   1  'Graphical
      TabIndex        =   38
      Tag             =   "MCTransferencias"
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton bSubrubroAcreedoresVarios 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":2E5E
      Style           =   1  'Graphical
      TabIndex        =   35
      Tag             =   "SubrubroAcreedoresVarios"
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton bSubrubroCompraMercaderia 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":3160
      Style           =   1  'Graphical
      TabIndex        =   32
      Tag             =   "SubrubroCompraMercaderia"
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton bSubrubroDeudoresPorVenta 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":3462
      Style           =   1  'Graphical
      TabIndex        =   25
      Tag             =   "SubrubroDeudoresPorVenta"
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton bSubrubroCobranzaVtasTel 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":3764
      Style           =   1  'Graphical
      TabIndex        =   24
      Tag             =   "SubrubroCobranzaVtasTel"
      Top             =   4200
      Width           =   375
   End
   Begin VB.CommandButton bSubrubroVtasTelACobrar 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":3A66
      Style           =   1  'Graphical
      TabIndex        =   23
      Tag             =   "SubrubroVtasTelACobrar"
      Top             =   4560
      Width           =   375
   End
   Begin VB.CommandButton bSubrubroCDAlCobro 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":3D68
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "SubrubroCDAlCobro"
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton bIngresosOperativos 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":406A
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "MCIngresosOperativos"
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton bpaMCVtaTelefonica 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":436C
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "MCVtaTelefonica"
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton bpaMCLiquidacionCamionero 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":466E
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "MCLiquidacionCamionero"
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton bpaMCNotaCredito 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":4970
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "MCNotaCredito"
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton bpaMCAnulacion 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":4C72
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "MCAnulacion"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton bpaMCChequeDiferido 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":4F74
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "MCChequeDiferido"
      Top             =   120
      Width           =   375
   End
   Begin AACombo99.AACombo cpaMCChequeDiferido 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
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
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7470
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   11165
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin AACombo99.AACombo cpaMCAnulacion 
      Height          =   315
      Left            =   2040
      TabIndex        =   5
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cpaMCNotaCredito 
      Height          =   315
      Left            =   2040
      TabIndex        =   8
      Top             =   840
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cpaMCLiquidacionCamionero 
      Height          =   315
      Left            =   2040
      TabIndex        =   11
      Top             =   1200
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cpaMCVtaTelefonica 
      Height          =   315
      Left            =   2040
      TabIndex        =   14
      Top             =   1560
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cIngresosOperativos 
      Height          =   315
      Left            =   2040
      TabIndex        =   17
      Top             =   1920
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cSubrubroCDAlCobro 
      Height          =   315
      Left            =   2040
      TabIndex        =   20
      Top             =   3480
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cSubrubroDeudoresPorVenta 
      Height          =   315
      Left            =   2040
      TabIndex        =   26
      Top             =   3840
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cSubrubroCobranzaVtasTel 
      Height          =   315
      Left            =   2040
      TabIndex        =   27
      Top             =   4200
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cSubrubroVtasTelACobrar 
      Height          =   315
      Left            =   2040
      TabIndex        =   28
      Top             =   4560
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cSubrubroCompraMercaderia 
      Height          =   315
      Left            =   2040
      TabIndex        =   33
      Top             =   5280
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cSubrubroAcreedoresVarios 
      Height          =   315
      Left            =   2040
      TabIndex        =   36
      Top             =   5640
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cMCTransferencias 
      Height          =   315
      Left            =   2040
      TabIndex        =   39
      Top             =   2280
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cSubrubroSeniasRecibidas 
      Height          =   315
      Left            =   2040
      TabIndex        =   42
      Top             =   6000
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cMCSenias 
      Height          =   315
      Left            =   2040
      TabIndex        =   45
      Top             =   2640
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cSubrubroVContado 
      Height          =   315
      Left            =   7440
      TabIndex        =   52
      Top             =   3480
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cSubrubroCCuotas 
      Height          =   315
      Left            =   7440
      TabIndex        =   53
      Top             =   3840
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cSubrubroCMorosidades 
      Height          =   315
      Left            =   7440
      TabIndex        =   54
      Top             =   4200
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cSubrubroNDevolucion 
      Height          =   315
      Left            =   7440
      TabIndex        =   55
      Top             =   4560
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cSubrubroNEspecial 
      Height          =   315
      Left            =   7440
      TabIndex        =   56
      Top             =   4920
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cSubRubroIngresosVarios 
      Height          =   315
      Left            =   7440
      TabIndex        =   64
      Top             =   5340
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cSubRubroIva 
      Height          =   315
      Left            =   7440
      TabIndex        =   65
      Top             =   5700
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cSubRubroVentas 
      Height          =   315
      Left            =   7440
      TabIndex        =   69
      Top             =   6060
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cSubrubroDifCambio 
      Height          =   315
      Left            =   2040
      TabIndex        =   72
      Top             =   6360
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cSubrubroDifCambioG 
      Height          =   315
      Left            =   2040
      TabIndex        =   75
      Top             =   6720
      Width           =   2895
      _ExtentX        =   5106
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
   Begin AACombo99.AACombo cSubrubroDifCostoImp 
      Height          =   315
      Left            =   2040
      TabIndex        =   78
      Top             =   7080
      Width           =   2895
      _ExtentX        =   5106
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
   Begin VB.Label Label25 
      Caption         =   "Dif. Costo Importación:"
      Height          =   255
      Left            =   120
      TabIndex        =   79
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label Label24 
      Caption         =   "Dif. de Cambio Ganada:"
      Height          =   255
      Left            =   120
      TabIndex        =   76
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Label Label23 
      Caption         =   "Dif. de Cambio Perdida:"
      Height          =   255
      Left            =   120
      TabIndex        =   73
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label22 
      Caption         =   "Ventas:"
      Height          =   255
      Left            =   5640
      TabIndex        =   70
      Top             =   6060
      Width           =   1695
   End
   Begin VB.Label Label21 
      Caption         =   "Ingresos Varios:"
      Height          =   255
      Left            =   5640
      TabIndex        =   67
      Top             =   5340
      Width           =   1695
   End
   Begin VB.Label Label20 
      Caption         =   "I.V.A. Ventas:"
      Height          =   255
      Left            =   5640
      TabIndex        =   66
      Top             =   5700
      Width           =   1695
   End
   Begin VB.Label Label19 
      Caption         =   "Ventas Contado:"
      Height          =   255
      Left            =   5640
      TabIndex        =   61
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label18 
      Caption         =   "Notas de Devolución:"
      Height          =   255
      Left            =   5640
      TabIndex        =   60
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label17 
      Caption         =   "Cob. de Morosidades:"
      Height          =   255
      Left            =   5640
      TabIndex        =   59
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label16 
      Caption         =   "Cobranza de Ctas.:"
      Height          =   255
      Left            =   5640
      TabIndex        =   58
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label15 
      Caption         =   "Notas Ctdo. Especial:"
      Height          =   255
      Left            =   5640
      TabIndex        =   57
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label14 
      Caption         =   "Señas:"
      Height          =   255
      Left            =   120
      TabIndex        =   46
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label13 
      Caption         =   "Señas (colectivos):"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label Label12 
      Caption         =   "Transferencias Ctas.:"
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "Acreedores Varios:"
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "Compra de Mercaderia:"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "N. Crédito / Anulaciones:"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Liq. de Camioneros:"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Ventas Telefónicas:"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Movimientos && Rubros "
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
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3120
      Width           =   10695
   End
   Begin VB.Label Label5 
      Caption         =   "Cheques Diferidos:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Ingresos Operativos:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Ventas Telefónicas:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Liq. de Camioneros:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Notas de Crédito:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lpaRubroDisponibilidad 
      Caption         =   "Anulaciones:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Tag             =   "paRubroDisponibilidad"
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lpaRubroImportaciones 
      Caption         =   "Cheques Diferidos:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bIngresosOperativos_Click()
    On Error Resume Next
    GraboParametro bIngresosOperativos.Tag, cIngresosOperativos.ItemData(cIngresosOperativos.ListIndex)
End Sub

Private Sub bMCSenias_Click()
    On Error Resume Next
    GraboParametro bMCSenias.Tag, cMCSenias.ItemData(cMCSenias.ListIndex)
End Sub

Private Sub bMCTransferencias_Click()
    On Error Resume Next
    GraboParametro bMCTransferencias.Tag, cMCTransferencias.ItemData(cMCTransferencias.ListIndex)
End Sub

Private Sub bpaMCAnulacion_Click()
    On Error Resume Next
    GraboParametro bpaMCAnulacion.Tag, cpaMCAnulacion.ItemData(cpaMCAnulacion.ListIndex)
End Sub

Private Sub bpaMCChequeDiferido_Click()
    On Error Resume Next
    GraboParametro bpaMCChequeDiferido.Tag, cpaMCChequeDiferido.ItemData(cpaMCChequeDiferido.ListIndex)
End Sub

Private Sub bpaMCLiquidacionCamionero_Click()
    On Error Resume Next
    GraboParametro bpaMCLiquidacionCamionero.Tag, cpaMCLiquidacionCamionero.ItemData(cpaMCLiquidacionCamionero.ListIndex)
End Sub

Private Sub bpaMCNotaCredito_Click()
    On Error Resume Next
    GraboParametro bpaMCNotaCredito.Tag, cpaMCNotaCredito.ItemData(cpaMCNotaCredito.ListIndex)
End Sub

Private Sub bpaMCVtaTelefonica_Click()
    On Error Resume Next
    GraboParametro bpaMCVtaTelefonica.Tag, cpaMCVtaTelefonica.ItemData(cpaMCVtaTelefonica.ListIndex)
End Sub

Private Sub bSubrubroAcreedoresVarios_Click()
    On Error Resume Next
    GraboParametro bSubrubroAcreedoresVarios.Tag, cSubrubroAcreedoresVarios.ItemData(cSubrubroAcreedoresVarios.ListIndex)
End Sub

Private Sub bSubrubroCCuotas_Click()
    On Error Resume Next
    GraboParametro bSubrubroCCuotas.Tag, cSubrubroCCuotas.ItemData(cSubrubroCCuotas.ListIndex)
End Sub

Private Sub bSubrubroCDAlCobro_Click()
    On Error Resume Next
    GraboParametro bSubrubroCDAlCobro.Tag, cSubrubroCDAlCobro.ItemData(cSubrubroCDAlCobro.ListIndex)
End Sub

Private Sub bSubrubroCMorosidades_Click()
    On Error Resume Next
    GraboParametro bSubrubroCMorosidades.Tag, cSubrubroCMorosidades.ItemData(cSubrubroCMorosidades.ListIndex)
End Sub

Private Sub bSubrubroCobranzaVtasTel_Click()
    On Error Resume Next
    GraboParametro bSubrubroCobranzaVtasTel.Tag, cSubrubroCobranzaVtasTel.ItemData(cSubrubroCobranzaVtasTel.ListIndex)
End Sub

Private Sub bSubrubroCompraMercaderia_Click()
    On Error Resume Next
    GraboParametro bSubrubroCompraMercaderia.Tag, cSubrubroCompraMercaderia.ItemData(cSubrubroCompraMercaderia.ListIndex)
End Sub

Private Sub bSubrubroDeudoresPorVenta_Click()
    On Error Resume Next
    GraboParametro bSubrubroDeudoresPorVenta.Tag, cSubrubroDeudoresPorVenta.ItemData(cSubrubroDeudoresPorVenta.ListIndex)
End Sub

Private Sub bSubrubroDifCambio_Click()
    On Error Resume Next
    GraboParametro bSubrubroDifCambio.Tag, cSubrubroDifCambio.ItemData(cSubrubroDifCambio.ListIndex)

End Sub

Private Sub bSubrubroDifCambioG_Click()
    On Error Resume Next
    GraboParametro bSubrubroDifCambioG.Tag, cSubrubroDifCambioG.ItemData(cSubrubroDifCambioG.ListIndex)
End Sub

Private Sub bSubrubroDifCostoImp_Click()
    On Error Resume Next
    GraboParametro bSubrubroDifCostoImp.Tag, cSubrubroDifCostoImp.ItemData(cSubrubroDifCostoImp.ListIndex)
End Sub

Private Sub bSubRubroIngresosVarios_Click()
    On Error Resume Next
    GraboParametro bSubRubroIngresosVarios.Tag, cSubRubroIngresosVarios.ItemData(cSubRubroIngresosVarios.ListIndex)
End Sub

Private Sub bSubRubroIva_Click()
    On Error Resume Next
    GraboParametro bSubRubroIva.Tag, cSubRubroIva.ItemData(cSubRubroIva.ListIndex)
End Sub

Private Sub bSubrubroNDevolucion_Click()
    On Error Resume Next
    GraboParametro bSubrubroNDevolucion.Tag, cSubrubroNDevolucion.ItemData(cSubrubroNDevolucion.ListIndex)
End Sub

Private Sub bSubrubroNEspecial_Click()
    On Error Resume Next
    GraboParametro bSubrubroNEspecial.Tag, cSubrubroNEspecial.ItemData(cSubrubroNEspecial.ListIndex)
End Sub

Private Sub bSubrubroSeniasRecibidas_Click()
    On Error Resume Next
    GraboParametro bSubrubroSeniasRecibidas.Tag, cSubrubroSeniasRecibidas.ItemData(cSubrubroSeniasRecibidas.ListIndex)
End Sub

Private Sub bSubrubroVContado_Click()
    On Error Resume Next
    GraboParametro bSubrubroVContado.Tag, cSubrubroVContado.ItemData(cSubrubroVContado.ListIndex)
End Sub

Private Sub bSubRubroVentas_Click()
    On Error Resume Next
    GraboParametro bSubRubroVentas.Tag, cSubRubroVentas.ItemData(cSubRubroVentas.ListIndex)
End Sub

Private Sub bSubrubroVtasTelACobrar_Click()
    On Error Resume Next
    GraboParametro bSubrubroVtasTelACobrar.Tag, cSubrubroVtasTelACobrar.ItemData(cSubrubroVtasTelACobrar.ListIndex)
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()

    Cons = "Select TMDCodigo, TMDNombre from TipoMovDisponibilidad Order by TMDNombre"
    CargoCombo Cons, cpaMCChequeDiferido: BuscoCodigoEnCombo cpaMCChequeDiferido, paMCChequeDiferido
    CargoCombo Cons, cpaMCAnulacion: BuscoCodigoEnCombo cpaMCAnulacion, paMCAnulacion
    CargoCombo Cons, cpaMCNotaCredito: BuscoCodigoEnCombo cpaMCNotaCredito, paMCNotaCredito
    CargoCombo Cons, cpaMCLiquidacionCamionero: BuscoCodigoEnCombo cpaMCLiquidacionCamionero, paMCLiquidacionCamionero
    CargoCombo Cons, cpaMCVtaTelefonica: BuscoCodigoEnCombo cpaMCVtaTelefonica, paMCVtaTelefonica
    CargoCombo Cons, cIngresosOperativos: BuscoCodigoEnCombo cIngresosOperativos, paMCIngresosOperativos
    CargoCombo Cons, cMCTransferencias: BuscoCodigoEnCombo cMCTransferencias, paMCTransferencias
    CargoCombo Cons, cMCSenias: BuscoCodigoEnCombo cMCSenias, paMCSenias
    
    Cons = "Select SRuID, SRuNombre from SubRubro Order by SRuNombre "
    CargoCombo Cons, cSubrubroVtasTelACobrar: BuscoCodigoEnCombo cSubrubroVtasTelACobrar, paSubrubroVtasTelACobrar
    CargoCombo Cons, cSubrubroDeudoresPorVenta: BuscoCodigoEnCombo cSubrubroDeudoresPorVenta, paSubrubroDeudoresPorVenta
    CargoCombo Cons, cSubrubroCobranzaVtasTel: BuscoCodigoEnCombo cSubrubroCobranzaVtasTel, paSubrubroCobranzaVtasTel
    CargoCombo Cons, cSubrubroCDAlCobro: BuscoCodigoEnCombo cSubrubroCDAlCobro, paSubrubroCDAlCobro
    CargoCombo Cons, cSubrubroCompraMercaderia: BuscoCodigoEnCombo cSubrubroCompraMercaderia, paSubrubroCompraMercaderia
    CargoCombo Cons, cSubrubroAcreedoresVarios: BuscoCodigoEnCombo cSubrubroAcreedoresVarios, paSubrubroAcreedoresVarios
    CargoCombo Cons, cSubrubroSeniasRecibidas: BuscoCodigoEnCombo cSubrubroSeniasRecibidas, paSubrubroSeniasRecibidas
    CargoCombo Cons, cSubrubroDifCambio: BuscoCodigoEnCombo cSubrubroDifCambio, paSubrubroDifCambio
    CargoCombo Cons, cSubrubroDifCambioG: BuscoCodigoEnCombo cSubrubroDifCambioG, paSubrubroDifCambioG
    CargoCombo Cons, cSubrubroDifCostoImp: BuscoCodigoEnCombo cSubrubroDifCostoImp, paSubrubroDifCostoImp
    
    CargoCombo Cons, cSubrubroNDevolucion: BuscoCodigoEnCombo cSubrubroNDevolucion, paSubrubroNDevolucion
    CargoCombo Cons, cSubrubroNEspecial: BuscoCodigoEnCombo cSubrubroNEspecial, paSubrubroNEspecial
    CargoCombo Cons, cSubrubroVContado: BuscoCodigoEnCombo cSubrubroVContado, paSubrubroVContado
    CargoCombo Cons, cSubrubroCCuotas: BuscoCodigoEnCombo cSubrubroCCuotas, paSubrubroCCuotas
    CargoCombo Cons, cSubrubroCMorosidades: BuscoCodigoEnCombo cSubrubroCMorosidades, paSubrubroCMorosidades
    
    CargoCombo Cons, cSubRubroIngresosVarios: BuscoCodigoEnCombo cSubRubroIngresosVarios, paSubrubroIngresosVarios
    CargoCombo Cons, cSubRubroIva: BuscoCodigoEnCombo cSubRubroIva, paSubrubroIVA
    CargoCombo Cons, cSubRubroVentas: BuscoCodigoEnCombo cSubRubroVentas, paSubrubroVentas
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    CierroConexion
    Set miConexion = Nothing
    Set msgError = Nothing
    Set clsGeneral = Nothing
    End
End Sub

Private Sub GraboParametro(Parametro As String, Valor As Variant)

    On Error GoTo errGrabar
    Screen.MousePointer = 11
    Parametro = Trim(Parametro)
    Cons = "Select * from Parametro Where ParNombre = '" & Parametro & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        RsAux.Edit
        RsAux!ParValor = Valor
        RsAux.Update: RsAux.Close
    Else
        If MsgBox("El parámetro " & Parametro & " no existe desea agregarlo.", vbQuestion + vbYesNo + vbDefaultButton2, "GRABAR PARÁMETRO") = vbYes Then
            RsAux.AddNew
            RsAux!ParNombre = Parametro
            RsAux!ParValor = Valor
            RsAux.Update
        End If
        RsAux.Close
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errGrabar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al grabar el parametro.", Err.Description
End Sub
