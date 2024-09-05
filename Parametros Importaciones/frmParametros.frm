VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.0#0"; "AACOMBO.OCX"
Begin VB.Form frmParametros 
   Caption         =   "Parámetros de Importaciones"
   ClientHeight    =   6630
   ClientLeft      =   2100
   ClientTop       =   2385
   ClientWidth     =   8985
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
   ScaleHeight     =   6630
   ScaleWidth      =   8985
   Begin VB.TextBox tPaRubroDisponibilidad 
      Height          =   285
      Left            =   2040
      TabIndex        =   63
      Text            =   "Text1"
      Top             =   480
      Width           =   1395
   End
   Begin VB.CommandButton bpaVencimientoLC 
      Height          =   315
      Left            =   8400
      Picture         =   "frmParametros.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   61
      Tag             =   "VencimientoLC"
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox tpaVencimientoLC 
      Height          =   285
      Left            =   7680
      TabIndex        =   60
      Text            =   "Text1"
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton bpaDisponibleArribo 
      Height          =   315
      Left            =   8400
      Picture         =   "frmParametros.frx":0744
      Style           =   1  'Graphical
      TabIndex        =   58
      Tag             =   "DisponibleArribo"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox tpaDisponibleArribo 
      Height          =   285
      Left            =   7680
      TabIndex        =   57
      Text            =   "Text1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton bpaDisponibleAArribar 
      Height          =   315
      Left            =   8400
      Picture         =   "frmParametros.frx":0A46
      Style           =   1  'Graphical
      TabIndex        =   55
      Tag             =   "DisponibleAArribar"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox tpaDisponibleAArribar 
      Height          =   285
      Left            =   7680
      TabIndex        =   54
      Text            =   "Text1"
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton bpaDisponibleAEmbarcar 
      Height          =   315
      Left            =   8400
      Picture         =   "frmParametros.frx":0D48
      Style           =   1  'Graphical
      TabIndex        =   52
      Tag             =   "DisponibleAEmbarcar"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox tpaDisponibleAEmbarcar 
      Height          =   285
      Left            =   7680
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton bpaDisponibleEnPuerto 
      Height          =   315
      Left            =   8400
      Picture         =   "frmParametros.frx":104A
      Style           =   1  'Graphical
      TabIndex        =   49
      Tag             =   "DisponibleEnPuerto"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox tpaDisponibleEnPuerto 
      Height          =   285
      Left            =   7680
      TabIndex        =   48
      Text            =   "Text1"
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton bpaDisponibleZF 
      Height          =   315
      Left            =   8400
      Picture         =   "frmParametros.frx":134C
      Style           =   1  'Graphical
      TabIndex        =   46
      Tag             =   "DisponibleZF"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox tpaDisponibleZF 
      Height          =   285
      Left            =   7680
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton bpaTipoTelefonoE 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":164E
      Style           =   1  'Graphical
      TabIndex        =   40
      Tag             =   "TipoTelefonoE"
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton bpaCategoriaCliente 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":1950
      Style           =   1  'Graphical
      TabIndex        =   39
      Tag             =   "CategoriaCliente"
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton bpaRepuesto 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":1C52
      Style           =   1  'Graphical
      TabIndex        =   36
      Tag             =   "Repuesto"
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton bpaMonedaPesos 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":1F54
      Style           =   1  'Graphical
      TabIndex        =   31
      Tag             =   "MonedaPesos"
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton bpaMonedaDolar 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":2256
      Style           =   1  'Graphical
      TabIndex        =   30
      Tag             =   "MonedaDolar"
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton bpaLocalZF 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":2558
      Style           =   1  'Graphical
      TabIndex        =   21
      Tag             =   "LocalZF"
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton bPaLocalPuerto 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":285A
      Style           =   1  'Graphical
      TabIndex        =   20
      Tag             =   "LocalPuerto"
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton bpaDepartamento 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":2B5C
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "Departamento"
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton bpaLocalidad 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":2E5E
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "Localidad"
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton bpaMDPagoDeCompra 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":3160
      Style           =   1  'Graphical
      TabIndex        =   15
      Tag             =   "MDPagoDeCompra"
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton bPaSubRubroTransporteM 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":3462
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "SubRubroTransporteM"
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton bPaSubRubroTransporteT 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":3764
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "SubRubroTransporteT"
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton bPaSubRubroDivisa 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":3A66
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "SubRubroDivisa"
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton BPaRubroDisponibilidad 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":3D68
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "RubroDisponibilidad"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton bPaRubroImportaciones 
      Height          =   315
      Left            =   5040
      Picture         =   "frmParametros.frx":406A
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "RubroImportaciones"
      Top             =   120
      Width           =   375
   End
   Begin AACombo99.AACombo cPaRubroImportaciones 
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
      Top             =   6375
      Width           =   8985
      _ExtentX        =   15849
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
            Object.Width           =   7646
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin AACombo99.AACombo cPaSubRubroDivisa 
      Height          =   315
      Left            =   2040
      TabIndex        =   7
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
   Begin AACombo99.AACombo cPaSubRubroTransporteT 
      Height          =   315
      Left            =   2040
      TabIndex        =   10
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
   Begin AACombo99.AACombo cPaSubRubroTransporteM 
      Height          =   315
      Left            =   2040
      TabIndex        =   13
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
   Begin AACombo99.AACombo cpaMDPagoDeCompra 
      Height          =   315
      Left            =   2040
      TabIndex        =   16
      Top             =   2160
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
   Begin AACombo99.AACombo cpaLocalZF 
      Height          =   315
      Left            =   2040
      TabIndex        =   22
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
   Begin AACombo99.AACombo cPaLocalPuerto 
      Height          =   315
      Left            =   2040
      TabIndex        =   23
      Top             =   3000
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
   Begin AACombo99.AACombo cpaDepartamento 
      Height          =   315
      Left            =   2040
      TabIndex        =   24
      Top             =   3360
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
   Begin AACombo99.AACombo cpaLocalidad 
      Height          =   315
      Left            =   2040
      TabIndex        =   25
      Top             =   3720
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
   Begin AACombo99.AACombo cpaMonedaPesos 
      Height          =   315
      Left            =   2040
      TabIndex        =   32
      Top             =   4320
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
   Begin AACombo99.AACombo cpaMonedaDolar 
      Height          =   315
      Left            =   2040
      TabIndex        =   33
      Top             =   4680
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
   Begin AACombo99.AACombo cpaRepuesto 
      Height          =   315
      Left            =   2040
      TabIndex        =   37
      Top             =   5160
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
   Begin AACombo99.AACombo cpaCategoriaCliente 
      Height          =   315
      Left            =   2040
      TabIndex        =   41
      Top             =   5520
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
   Begin AACombo99.AACombo cpaTipoTelefonoE 
      Height          =   315
      Left            =   2040
      TabIndex        =   42
      Top             =   5880
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
   Begin VB.Label Label19 
      Caption         =   "Vencimiento LC:"
      Height          =   255
      Left            =   5880
      TabIndex        =   62
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label18 
      Caption         =   "Disponible Arribo:"
      Height          =   255
      Left            =   5880
      TabIndex        =   59
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label17 
      Caption         =   "Disponible A Arribar:"
      Height          =   255
      Left            =   5880
      TabIndex        =   56
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label16 
      Caption         =   "Disponible A Embarcar:"
      Height          =   255
      Left            =   5880
      TabIndex        =   53
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label15 
      Caption         =   "Disponible Puerto:"
      Height          =   255
      Left            =   5880
      TabIndex        =   50
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label14 
      Caption         =   "Disponible Zona Franca:"
      Height          =   255
      Left            =   5880
      TabIndex        =   47
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label13 
      Caption         =   "Categoría Cliente:"
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label12 
      Caption         =   "Telefono Empresa:"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "Grupo Repuesto:"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "Moneda Dolar:"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Moneda Pesos:"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Local Zona Franca:"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Tag             =   "paRubroDisponibilidad"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Local Puerto:"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Departamento:"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Localidad:"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Mov. Disp. Pago Compras:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "SubR. Transp. Marítim:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "SubR. Transp. Terres:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "SubRubro Divisa:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lpaRubroDisponibilidad 
      Caption         =   "Rubro Disponibilidad:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Tag             =   "paRubroDisponibilidad"
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lpaRubroImportaciones 
      Caption         =   "Rubro Importaciones:"
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
Private Sub bpaCategoriaCliente_Click()
    GraboParametro bpaCategoriaCliente.Tag, cpaCategoriaCliente.ItemData(cpaCategoriaCliente.ListIndex)
End Sub

Private Sub bpaDepartamento_Click()
    GraboParametro bpaDepartamento.Tag, cpaDepartamento.ItemData(cpaDepartamento.ListIndex)
End Sub

Private Sub bpaDisponibleAArribar_Click()
    GraboParametro bpaDisponibleAArribar.Tag, tpaDisponibleAArribar.Text
End Sub

Private Sub bpaDisponibleAEmbarcar_Click()
    GraboParametro bpaDisponibleAEmbarcar.Tag, tpaDisponibleAEmbarcar.Text
End Sub

Private Sub bpaDisponibleArribo_Click()
    GraboParametro bpaDisponibleArribo.Tag, tpaDisponibleArribo.Text
End Sub

Private Sub bpaDisponibleEnPuerto_Click()
    GraboParametro bpaDisponibleEnPuerto.Tag, tpaDisponibleEnPuerto.Text
End Sub

Private Sub bpaDisponibleZF_Click()
    GraboParametro bpaDisponibleZF.Tag, tpaDisponibleZF.Text
End Sub

Private Sub bpaLocalidad_Click()
    GraboParametro bpaLocalidad.Tag, cpaLocalidad.ItemData(cpaLocalidad.ListIndex)
End Sub

Private Sub bPaLocalPuerto_Click()
    GraboParametro bPaLocalPuerto.Tag, cPaLocalPuerto.ItemData(cPaLocalPuerto.ListIndex)
End Sub

Private Sub bpaLocalZF_Click()
    GraboParametro bpaLocalZF.Tag, cpaLocalZF.ItemData(cpaLocalZF.ListIndex)
End Sub

Private Sub bpaMDPagoDeCompra_Click()
    GraboParametro bpaMDPagoDeCompra.Tag, cpaMDPagoDeCompra.ItemData(cpaMDPagoDeCompra.ListIndex)
End Sub

Private Sub bpaMonedaDolar_Click()
    GraboParametro bpaMonedaDolar.Tag, cpaMonedaDolar.ItemData(cpaMonedaDolar.ListIndex)
End Sub

Private Sub bpaMonedaPesos_Click()
    GraboParametro bpaMonedaPesos.Tag, cpaMonedaPesos.ItemData(cpaMonedaPesos.ListIndex)
End Sub

Private Sub bpaRepuesto_Click()
    GraboParametro bpaRepuesto.Tag, cpaRepuesto.ItemData(cpaRepuesto.ListIndex)
End Sub

Private Sub BPaRubroDisponibilidad_Click()
    If Not IsNumeric(tPaRubroDisponibilidad.Text) Then
        MsgBox "Debe ingresar los primeros números del rubro disponibilidades.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    GraboParametro BPaRubroDisponibilidad.Tag, tPaRubroDisponibilidad.Text
End Sub

Private Sub bPaRubroImportaciones_Click()
    GraboParametro bPaRubroImportaciones.Tag, cPaRubroImportaciones.ItemData(cPaRubroImportaciones.ListIndex)
End Sub

Private Sub bPaSubRubroDivisa_Click()
    GraboParametro bPaSubRubroDivisa.Tag, cPaSubRubroDivisa.ItemData(cPaSubRubroDivisa.ListIndex)
End Sub

Private Sub bPaSubRubroTransporteM_Click()
    GraboParametro bPaSubRubroTransporteM.Tag, cPaSubRubroTransporteM.ItemData(cPaSubRubroTransporteM.ListIndex)
End Sub

Private Sub bPaSubRubroTransporteT_Click()
    GraboParametro bPaSubRubroTransporteT.Tag, cPaSubRubroTransporteT.ItemData(cPaSubRubroTransporteT.ListIndex)
End Sub

Private Sub bpaTipoTelefonoE_Click()
    GraboParametro bpaTipoTelefonoE.Tag, cpaTipoTelefonoE.ItemData(cpaTipoTelefonoE.ListIndex)
End Sub

Private Sub bpaVencimientoLC_Click()
    GraboParametro bpaVencimientoLC.Tag, tpaVencimientoLC.Text
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()

'Parametros.------------------------------------------------------
'Public paEstadoArticuloEntrega As Long

    Cons = "Select RubID, RubNombre from Rubro Order by RubNombre"
    CargoCombo Cons, cPaRubroImportaciones: BuscoCodigoEnCombo cPaRubroImportaciones, paRubroImportaciones
    tPaRubroDisponibilidad.Text = paRubroDisponibilidad
    
    Cons = "Select SRuID, SRuNombre from SubRubro Order by SRuNombre"
    CargoCombo Cons, cPaSubRubroDivisa: BuscoCodigoEnCombo cPaSubRubroDivisa, paSubrubroDivisa
    CargoCombo Cons, cPaSubRubroTransporteT: BuscoCodigoEnCombo cPaSubRubroTransporteT, paSubrubroTransporteT
    CargoCombo Cons, cPaSubRubroTransporteM: BuscoCodigoEnCombo cPaSubRubroTransporteM, paSubrubroTransporteM
    
    Cons = "Select TMDCodigo, TMDNombre from TipoMovDisponibilidad Order by TMDNombre"
    CargoCombo Cons, cpaMDPagoDeCompra: BuscoCodigoEnCombo cpaMDPagoDeCompra, paMDPagoDeCompra
    
    Cons = "Select LocCodigo, LocNombre from Local Order by LocNombre"
    CargoCombo Cons, cpaLocalZF: BuscoCodigoEnCombo cpaLocalZF, CLng(paLocalZF)
    CargoCombo Cons, cPaLocalPuerto: BuscoCodigoEnCombo cPaLocalPuerto, CLng(paLocalPuerto)
    
    Cons = "Select DepCodigo, DepNombre from Departamento Order by DepNombre"
    CargoCombo Cons, cpaDepartamento: BuscoCodigoEnCombo cpaDepartamento, paDepartamento
    
    Cons = "Select LocCodigo, LocNombre from Localidad Order by LocNombre"
    CargoCombo Cons, cpaLocalidad: BuscoCodigoEnCombo cpaLocalidad, paLocalidad
    
    Cons = "Select MonCodigo, MonSigno from Moneda "
    CargoCombo Cons, cpaMonedaPesos: BuscoCodigoEnCombo cpaMonedaPesos, CLng(paMonedaPesos)
    CargoCombo Cons, cpaMonedaDolar: BuscoCodigoEnCombo cpaMonedaDolar, CLng(paMonedaDolar)
    
    Cons = "Select GruCodigo, GruNombre from Grupo Order by GruNombre"
    CargoCombo Cons, cpaRepuesto: BuscoCodigoEnCombo cpaRepuesto, paRepuesto
    
    Cons = "Select TTeCodigo, TTeNombre from TipoTelefono Order by TTeNombre"
    CargoCombo Cons, cpaTipoTelefonoE: BuscoCodigoEnCombo cpaTipoTelefonoE, paTipoTelefonoE
    
    Cons = "Select CClCodigo, CClNombre from CategoriaCliente Order by CClNombre"
    CargoCombo Cons, cpaCategoriaCliente: BuscoCodigoEnCombo cpaCategoriaCliente, paCategoriaCliente
        
    tpaDisponibleZF.Text = paDisponibleZF
    tpaDisponibleArribo.Text = paDisponibleArribo
    tpaDisponibleAArribar.Text = paDisponibleAArribar
    tpaDisponibleAEmbarcar.Text = paDisponibleAEmbarcar
    tpaDisponibleEnPuerto.Text = paDisponibleEnPuerto
    tpaVencimientoLC.Text = paVencimientoLC
    
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
