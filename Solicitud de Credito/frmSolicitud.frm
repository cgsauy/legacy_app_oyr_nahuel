VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{1292AE18-2B08-4CE3-9F79-9CB06F26AB54}#1.7#0"; "orEMails.ocx"
Object = "{5EA2D00A-68AC-4888-98E6-53F6035BBEE3}#1.2#0"; "CGSABuscarCliente.ocx"
Begin VB.Form frmSolicitud 
   BackColor       =   &H00C0B000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitudes de Compra"
   ClientHeight    =   5550
   ClientLeft      =   2025
   ClientTop       =   3195
   ClientWidth     =   9465
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSolicitud.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   9465
   Begin VB.Timer tmArticuloLimitado 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   2160
      Top             =   120
   End
   Begin prjBuscarCliente.ucBuscarCliente txtGarantia 
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   1740
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      Text            =   "_.___.___-_"
      DocumentoCliente=   1
      QueryFind       =   "EXEC [dbo].[prg_BuscarCliente] 0, '', '', '', '', '', '[KeyQuery]', 0, 0, '', '', 7"
      KeyQuery        =   "[KeyQuery]"
      Comportamiento  =   2
   End
   Begin prjBuscarCliente.ucBuscarCliente txtCliente 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   540
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      Text            =   "_.___.___-_"
      DocumentoCliente=   1
      QueryFind       =   "EXEC [dbo].[prg_BuscarCliente] 0, '', '', '', '', '', '[KeyQuery]', 0, 0, '', '', 7"
      KeyQuery        =   "[KeyQuery]"
      NeedCheckDigit  =   0   'False
      Comportamiento  =   2
   End
   Begin VB.Timer tmClose 
      Enabled         =   0   'False
      Left            =   720
      Top             =   120
   End
   Begin orEMails.ctrEMails cEMailsT 
      Height          =   315
      Left            =   700
      TabIndex        =   5
      Top             =   1140
      Width           =   5175
      _ExtentX        =   5980
      _ExtentY        =   556
      ForeColor       =   16777215
   End
   Begin VB.ComboBox cTelsT 
      BackColor       =   &H000000C0&
      Height          =   315
      Left            =   6840
      Style           =   2  'Dropdown List
      TabIndex        =   53
      Top             =   1140
      Width           =   2415
   End
   Begin VB.ComboBox cMoneda 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   8460
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   60
      Width           =   855
   End
   Begin VB.ComboBox cDireccion 
      Height          =   315
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   825
      Width           =   1515
   End
   Begin AACombo99.AACombo cComentario 
      Height          =   315
      Left            =   1080
      TabIndex        =   29
      Top             =   4800
      Width           =   5775
      _ExtentX        =   10186
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
   Begin VB.TextBox tVendedor 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   5380
      MaxLength       =   3
      TabIndex        =   25
      Top             =   4455
      Width           =   375
   End
   Begin VB.TextBox tEntregaT 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1080
      MaxLength       =   12
      TabIndex        =   21
      Top             =   4455
      Width           =   1095
   End
   Begin VB.TextBox tEntrega 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   6960
      MaxLength       =   12
      TabIndex        =   17
      Top             =   2460
      Width           =   1215
   End
   Begin VB.TextBox tUsuario 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   6480
      MaxLength       =   3
      TabIndex        =   27
      Top             =   4455
      Width           =   375
   End
   Begin VB.ComboBox cArticulo 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1320
      Style           =   1  'Simple Combo
      TabIndex        =   11
      Top             =   2460
      Width           =   3855
   End
   Begin VB.TextBox tCantidad 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   5160
      MaxLength       =   5
      TabIndex        =   13
      Top             =   2460
      Width           =   615
   End
   Begin VB.TextBox tUnitario 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   5760
      MaxLength       =   12
      TabIndex        =   15
      Top             =   2460
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   32
      Top             =   5295
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11829
            MinWidth        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1799
            MinWidth        =   1482
            Text            =   "F2-Modificar "
            TextSave        =   "F2-Modificar "
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1482
            MinWidth        =   1482
            Text            =   "F3-Nuevo"
            TextSave        =   "F3-Nuevo"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1482
            MinWidth        =   1482
            Text            =   "F4-Buscar"
            TextSave        =   "F4-Buscar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvVenta 
      Height          =   1575
      Left            =   120
      TabIndex        =   19
      Top             =   2820
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Plan"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cant."
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Artículo"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Contado x1"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "I.V.A."
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Entrega"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Cuota"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Sub Total"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Financiado x1"
         Object.Width           =   0
      EndProperty
   End
   Begin AACombo99.AACombo cCuota 
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   2460
      Width           =   1215
      _ExtentX        =   2143
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
      Text            =   ""
   End
   Begin AACombo99.AACombo cPago 
      Height          =   315
      Left            =   2820
      TabIndex        =   23
      Top             =   4440
      Width           =   1635
      _ExtentX        =   2884
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
      Text            =   ""
   End
   Begin VB.Label lTelsT 
      BackStyle       =   0  'Transparent
      Caption         =   "Tels:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   52
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lRucCliente 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "90"
      ForeColor       =   &H80000008&
      Height          =   260
      Left            =   960
      TabIndex        =   51
      Top             =   850
      UseMnemonic     =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&eMails"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1220
      Width           =   555
   End
   Begin VB.Label lUnitario 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cuota x&1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5760
      TabIndex        =   14
      Top             =   2235
      Width           =   1215
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "&Vendedor:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4560
      TabIndex        =   24
      Top             =   4500
      Width           =   855
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&L"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2940
      Width           =   1215
   End
   Begin VB.Label lTEdad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "90"
      ForeColor       =   &H80000008&
      Height          =   265
      Left            =   8925
      TabIndex        =   50
      Top             =   540
      UseMnemonic     =   0   'False
      Width           =   315
   End
   Begin VB.Label lGEdad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "90"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   8925
      TabIndex        =   49
      Top             =   1740
      UseMnemonic     =   0   'False
      Width           =   315
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Edad:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8475
      TabIndex        =   48
      Top             =   550
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Edad:"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   8475
      TabIndex        =   47
      Top             =   1755
      Width           =   495
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Pag&o:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      TabIndex        =   22
      Top             =   4500
      Width           =   495
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Co&mentarios:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   105
      TabIndex        =   28
      Top             =   4875
      Width           =   975
   End
   Begin VB.Label lGarantia 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3360
      TabIndex        =   46
      Top             =   1740
      UseMnemonic     =   0   'False
      Width           =   4935
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2640
      TabIndex        =   45
      Top             =   1755
      Width           =   615
   End
   Begin VB.Label lblInfoAval 
      BackStyle       =   0  'Transparent
      Caption         =   "&Garantía:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1755
      Width           =   735
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "E&ntrega:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   105
      TabIndex        =   20
      Top             =   4500
      Width           =   735
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Entrega &Inic."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6960
      TabIndex        =   16
      Top             =   2235
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   9480
      Y1              =   5220
      Y2              =   5220
   End
   Begin VB.Label lSubTotalF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   8160
      TabIndex        =   44
      Top             =   2460
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sub Total (F)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8145
      TabIndex        =   43
      Top             =   2235
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   9480
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5820
      TabIndex        =   26
      Top             =   4500
      Width           =   735
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   6840
      TabIndex        =   42
      Top             =   4870
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "I.V.A.:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   6840
      TabIndex        =   41
      Top             =   4630
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SubTotal:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   6960
      TabIndex        =   40
      Top             =   4390
      Width           =   975
   End
   Begin VB.Label labSubTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6960
      TabIndex        =   39
      Top             =   4380
      Width           =   2415
   End
   Begin VB.Label labIVA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6960
      TabIndex        =   38
      Top             =   4620
      Width           =   2415
   End
   Begin VB.Label labTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6960
      TabIndex        =   37
      Top             =   4860
      Width           =   2415
   End
   Begin VB.Label lArticulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Artículo"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   2235
      Width           =   3855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Can&t."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5160
      TabIndex        =   12
      Top             =   2235
      Width           =   615
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "&Moneda:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7680
      TabIndex        =   30
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lbltitCliente 
      BackStyle       =   0  'Transparent
      Caption         =   "&Cliente:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   555
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "R.U.C.:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   900
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   36
      Top             =   550
      Width           =   855
   End
   Begin VB.Label lblNombreCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "Rodriguez Fernandez, Rodrigo Bernardino"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3360
      TabIndex        =   35
      Top             =   540
      UseMnemonic     =   0   'False
      Width           =   4935
   End
   Begin VB.Label lNDireccion 
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      Height          =   255
      Left            =   3000
      TabIndex        =   34
      Top             =   840
      Width           =   855
   End
   Begin VB.Label labDireccion 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "Niagara 2345"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4170
      TabIndex        =   33
      Top             =   840
      UseMnemonic     =   0   'False
      Width           =   5085
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   1095
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   435
      Width           =   9375
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "       &Plan"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2220
      Width           =   9255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   465
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   1635
      Width           =   9375
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuEmitir 
         Caption         =   "&Grabar Solicitud"
         Enabled         =   0   'False
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuVisulizacionOp 
         Caption         =   "Visualización de Operaciones"
         Shortcut        =   {F12}
      End
      Begin VB.Menu MnuLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuLimpiar 
         Caption         =   "&Limpiar Ficha"
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
   Begin VB.Menu MnuHelp 
      Caption         =   "&?"
      Begin VB.Menu MnuHlp 
         Caption         =   "Ayuda ..."
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu MnuMoussePersona 
      Caption         =   "&MoussePersona"
      Visible         =   0   'False
      Begin VB.Menu MnuMoCliente 
         Caption         =   "Menú Cliente"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuLineaMP2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPEmpleo 
         Caption         =   "Ingresar &Empleos"
      End
      Begin VB.Menu MnuPReferencia 
         Caption         =   "Ingresar &Referencias"
      End
      Begin VB.Menu MnuPTitulo 
         Caption         =   "Ingresar &Títulos"
      End
      Begin VB.Menu MnuLineaMP3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCancelarMP 
         Caption         =   "&Cancelar"
      End
   End
End
Attribute VB_Name = "frmSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oArtEdicion As clsArticulo
Dim colArtsGrilla As Collection

Dim bSeñalInhabilitadoXMayor As Byte

Public prmIDCliente As Long
Public prmIDLlamada As Long

Private Type typCombo
    Articulo As Long
    Q As Integer
    Bonificacion As Currency
    EsBonificacion As Boolean
End Type

Private Type typSucDP
    IDArticulo As Long
    TSuceso As Integer
    DifPrecio As Currency
    TextoPlan As String * 15
End Type

Dim arrSucDP() As typSucDP

Const FormatoCedula = "_.___.___-_"

Enum TFactura
    Articulo = 0
    Servicio = 1
    ArtEspecifico = 2
End Enum

Dim itmX As ListItem
Dim aTexto As String

Dim sConEntrega As Boolean
Dim sDistribuir As Boolean          'Para redistribuir las entregas

Private RsBoleta As rdoResultset
Private RsArt As rdoResultset

Dim gSucesoUsr As Long, gSucesoDef As String, gSucesoUsrAut As Long
Dim gDirFactura As Long         'Direccion con la que factura el cliente

Dim gConyugeDelGarante As Long
Dim mMRound As String                   'Patron para redondeo de la Moneda seleccioanda

Private Function ValidarTipoCuota() As Boolean
    
    If Trim(cCuota.Text) = "" Then
        If tEntregaT.Enabled Then tEntregaT.SetFocus Else cPago.SetFocus
        ValidarTipoCuota = True: Exit Function
    Else
        If cCuota.ListIndex <> -1 Then
            cArticulo_KeyPress (13)
            cArticulo.SetFocus: ValidarTipoCuota = True: Exit Function
        End If
        
        'Veo planes deshabilitados
        If MsgBox("El plan " & Trim(cCuota.Text) & " no está habilitado." & vbCrLf & _
                        "Desea continuar ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Plan No Habilitado") = vbNo Then Exit Function
        
        Screen.MousePointer = 11
        
        Dim bHay As Boolean
        Cons = " Select TCuCodigo, TCuAbreviacion as 'Tipo de Cuota', TCuNombre as 'Detalle'" & _
                   " From TipoCuota " & _
                   " Where TCuDeshabilitado = 'S' " & _
                   " Order by TCuNombre "
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then bHay = True Else bHay = False
        RsAux.Close
        
        Dim aSValor As Long, aSTexto As String
        aSValor = 0
        If Not bHay Then
            MsgBox "No hay tipos de cuotas deshabilitados.", vbExclamation, "No hay datos"
        Else
            Dim miLista As New clsListadeAyuda
            aSValor = miLista.ActivarAyuda(cBase, Cons, , 1, "Planes Deshabilitados")
            Me.Refresh
            If aSValor > 0 Then
                aSValor = miLista.RetornoDatoSeleccionado(0)
                aSTexto = Trim(miLista.RetornoDatoSeleccionado(1))
            End If
            Set miLista = Nothing
        End If
                    
        If aSValor <> 0 Then
            cCuota.AddItem aSTexto
            cCuota.ItemData(cCuota.NewIndex) = aSValor
            cCuota.Text = aSTexto
        End If
        Screen.MousePointer = 0
        ValidarTipoCuota = False 'Me quedo parado en la misma hasta que me de enter.
    End If
End Function

Private Sub cArticulo_GotFocus()
    
    cArticulo.SelStart = 0: cArticulo.SelLength = Len(cArticulo.Text)
    Select Case Val(lArticulo.Tag)
        Case TFactura.Servicio: Status.Panels(1).Text = "Ingrese el código de servicio que desea facturar."
        Case TFactura.Articulo: Status.Panels(1).Text = "Ingrese el código o nombre del artículo."
        Case TFactura.ArtEspecifico: Status.Panels(1).Text = "Ingrese el artículo específico a buscar."
    End Select
    
End Sub

Private Sub cArticulo_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF1:
            Select Case Val(lArticulo.Tag)
                Case TFactura.Articulo: TipoFacturacion TFactura.ArtEspecifico
                Case TFactura.ArtEspecifico: TipoFacturacion TFactura.Servicio
                Case TFactura.Servicio: TipoFacturacion TFactura.Articulo
            End Select
            Call cArticulo_GotFocus
    End Select
    
End Sub

Private Sub TipoFacturacion(Tipo As Byte)

    cArticulo.Clear: cArticulo.Tag = ""
    Select Case Tipo
        Case TFactura.Articulo
            lArticulo.Tag = TFactura.Articulo: lArticulo.Caption = "&Artículo"
            tCantidad.Enabled = True: tCantidad.BackColor = Obligatorio
            tUnitario.Enabled = True: tUnitario.BackColor = Obligatorio
    
        Case TFactura.ArtEspecifico
            lArticulo.Tag = TFactura.ArtEspecifico: lArticulo.Caption = "&Artículo Específico"
            tCantidad.Enabled = True: tCantidad.BackColor = Obligatorio
            tUnitario.Enabled = True: tUnitario.BackColor = Obligatorio
    
        Case TFactura.Servicio
            lArticulo.Tag = TFactura.Servicio: lArticulo.Caption = "&Servicio"
            tCantidad.Enabled = False: tCantidad.BackColor = Inactivo
            tUnitario.Enabled = False: tUnitario.BackColor = Inactivo
    End Select
    
End Sub

Private Sub cArticulo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Select Case Val(lArticulo.Tag)
        
            Case TFactura.Articulo:
                If Trim(cArticulo.Text) <> vbNullString Then
                    If Not ValidoBuscarArticulo Then Exit Sub
                    
                    If Not IsNumeric(cArticulo.Text) Then BuscoArticuloXNombre Else BuscoArticuloxCodigo CLng(cArticulo.Text), 0
                    
                    If cArticulo.ListIndex > -1 Then
                        If Trim(tCantidad.Text) = "" Then tCantidad.Text = "1"
                        Foco tCantidad
                    End If
                Else
                    If tEntregaT.Enabled Then tEntregaT.SetFocus Else cPago.SetFocus
                End If
                        
            Case TFactura.Servicio: If IsNumeric(cArticulo.Text) Then CargoDatosServicio CLng(cArticulo.Text)
            
            Case TFactura.ArtEspecifico
                If Trim(cArticulo.Text) <> "" Then
                    If Not ValidoBuscarArticulo Then Exit Sub
                    BuscoArticuloEspecifico cArticulo.Text
                    If cArticulo.ListIndex > -1 Then
                        tCantidad.Text = 1
                        Foco tUnitario
                    End If
                Else
                    If lvVenta.ListItems.Count > 0 Then
                        If tEntregaT.Enabled Then tEntregaT.SetFocus Else cPago.SetFocus
                    End If
                End If
                
        End Select
    End If

End Sub

Private Sub cComentario_LostFocus()
    cComentario.SelLength = 0
End Sub

Private Sub cCuota_Change()
    If Val(lArticulo.Tag) = TFactura.Servicio Or cArticulo.ListIndex = -1 Then
        LimpioRenglon
        If cCuota.ListIndex <> -1 Then Call cCuota_Click
    End If
End Sub

Private Sub cCuota_Click()

    On Error GoTo errBuscar
    If Val(lArticulo.Tag) = TFactura.Servicio Or cArticulo.ListIndex = -1 Then
        LimpioRenglon
    End If
    If cCuota.ListIndex = -1 Then Exit Sub
    
    Screen.MousePointer = 11
    
    'Busco los datos del tipo de cuota seleccionado
    Cons = "Select * from TipoCuota where TCuCodigo = " & cCuota.ItemData(cCuota.ListIndex)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not IsNull(RsAux!TCuVencimientoE) Then
        sConEntrega = True
        lUnitario.Caption = "Contado &x1"
    Else
        sConEntrega = False
        lUnitario.Caption = "Cuota &x1"
    End If
    cCuota.Tag = RsAux!TCuCantidad
    
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub

errBuscar:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el tipo de cuota.", Err.Description
    Screen.MousePointer = 0
    cCuota.Text = ""
End Sub

Private Sub cCuota_GotFocus()
    cCuota.SelStart = 0: cCuota.SelLength = Len(cCuota.Text)
    Status.Panels(1).Text = "Seleccione la financiación del artículo."
End Sub

Private Sub cCuota_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        ValidarTipoCuota
    End If

    Exit Sub
    
errBPlan:
    clsGeneral.OcurrioError "Error al buscar los planes no habilitados.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub cCuota_LostFocus()
    cCuota.SelLength = 0
End Sub

Private Sub cCuota_Validate(Cancel As Boolean)
    Cancel = Not ValidarTipoCuota()
End Sub

Private Sub cDireccion_Click()
On Error GoTo errCargar

    If cDireccion.ListIndex <> -1 Then
        Screen.MousePointer = 11
        labDireccion.Caption = ""
        labDireccion.Caption = " " & clsGeneral.ArmoDireccionEnTexto(cBase, cDireccion.ItemData(cDireccion.ListIndex))
        Screen.MousePointer = 0
    End If

errCargar:
    Screen.MousePointer = 0
End Sub

Private Sub cDireccion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cEMailsT.SetFocus
End Sub

Private Sub cEMailsT_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtGarantia.SetFocus
End Sub

Private Sub cMoneda_Click()
On Error Resume Next
    LimpioRenglon
    
    'Cargo variables para la moneda seleccioada
    mMRound = dis_arrMonedaProp(cMoneda.ItemData(cMoneda.ListIndex), enuMoneda.pRedondeo)
    
End Sub

Private Sub cMoneda_GotFocus()
    Status.Panels(1).Text = "Seleccione una moneda para facturar la solicitud."
End Sub

Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtCliente.SetFocus
End Sub

Private Sub cPago_GotFocus()
    cPago.SelStart = 0: cPago.SelLength = Len(cPago.Text)
    Status.Panels(1).Text = "Seleccione la forma de pago de la solicitud."
End Sub

Private Sub cPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tVendedor
End Sub

Private Sub cPago_LostFocus()
    cPago.SelLength = 0
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF2
            If Me.ActiveControl.Name <> txtGarantia.Name And Me.ActiveControl.Name <> txtCliente.Name Then txtCliente.EditarCliente
        Case 93
            If Me.ActiveControl.Name = txtCliente.Name Then
                If txtCliente.DocumentoCliente <> DC_RUT Then
                    MnuPEmpleo.Enabled = True
                    MnuPReferencia.Enabled = True
                    MnuPTitulo.Enabled = True
                    PopupMenu MnuMoussePersona, , txtCliente.Left + (txtCliente.Width / 2), (txtCliente.Top + txtCliente.Height) - (txtCliente.Height / 2)
                End If
            End If
        Case vbKeyC
            If Shift = vbAltMask And txtCliente.Enabled Then txtCliente.SetFocus
    End Select
    
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    Set txtCliente.Connect = cBase
    Set txtGarantia.Connect = cBase
    zfn_InicializoControles
    LimpioTodaLaFicha False
    If prmIDCliente <> 0 Then txtCliente.CargarControl prmIDCliente
        
End Sub

'------------------------------------------------------------------------------------------------
'   Carga los tipos de cuotas segun la categoría del cliente (Especiales o Normales)
'------------------------------------------------------------------------------------------------
Private Sub CargoCuotas(CategoriaCliente As Long)
    
    If CategoriaCliente = paCategoriaCliente Then
        Cons = "Select TCuCodigo, TCuAbreviacion From TipoCuota" _
                & " Where TCuCodigo <> " & paTipoCuotaContado _
                & " And TCuDeshabilitado = Null" _
                & " And TCuEspecial = 0" _
                & " Order by TCuAbreviacion"
    Else
        Cons = "Select TCuCodigo, TCuAbreviacion From TipoCuota" _
                & " Where TCuCodigo <> " & paTipoCuotaContado _
                & " And TCuDeshabilitado = Null" _
                & " Order by TCuAbreviacion"
    End If
    
    CargoCombo Cons, cCuota, ""

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Status.Panels(1).Text = vbNullString
End Sub

Private Sub Form_Resize()
On Error GoTo errRsz
    Line2.Y1 = Me.ScaleHeight - Status.Height - 10
    Line2.Y2 = Line2.Y1
errRsz:
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    GuardoSeteoForm Me
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    'Set oResAuto = Nothing
    End

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label12_Click()
    Foco cMoneda
End Sub

Private Sub Label13_Click()

    Foco tUsuario
    Status.Panels(1).Text = " Ingrese el dígito de usuario."

End Sub

Private Sub Label14_Click()
    Foco tEntrega
End Sub

Private Sub Label16_Click()
    Foco txtGarantia
End Sub

Private Sub Label17_Click()
    Foco cCuota
End Sub

Private Sub Label26_Click()
    Foco cComentario
End Sub

Private Sub Label27_Click()
    Foco cPago
End Sub

Private Sub Label30_Click()
    Foco tVendedor
End Sub


Private Sub Label4_Click()
    On Error Resume Next
    cEMailsT.SetFocus
End Sub

Private Sub lArticulo_Click()
    Foco cArticulo
End Sub

Private Sub Label6_Click()
    Foco tCantidad
End Sub

Private Sub lvVenta_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo ErrlvKD

    If lvVenta.ListItems.Count > 0 Then
        Select Case KeyCode
            Case vbKeySpace    'EDITO EL RENGLON----------------------------------------------------------------------------------
                If Val(lArticulo.Tag) <> TFactura.Articulo And Val(lArticulo.Tag) <> TFactura.ArtEspecifico Then Exit Sub
                
                If ValorClave(lvVenta.SelectedItem.Key, "F") > 0 Then TipoFacturacion Tipo:=TFactura.ArtEspecifico Else TipoFacturacion Tipo:=TFactura.Articulo
                
                lSubTotalF.Caption = lvVenta.SelectedItem.SubItems(7)   'Subtotal
                lSubTotalF.Tag = lvVenta.SelectedItem.SubItems(8)       'P.U. Financiado
                
                If sConEntrega Then
                    If Trim(lSubTotalF.Caption) = "" Then       'VALIDO QUE SE HAYA REALIZADO LA DISTRIBUCION
                        MsgBox "Aún no ha ingresado el valor de entrega total de la solicitud o no se realizó la distribución automática.", vbExclamation, "ATENCIÓN"
                        LimpioRenglon
                        If tEntregaT.Enabled Then tEntregaT.SetFocus
                        lSubTotalF.Caption = ""
                        lSubTotalF.Tag = ""
                        Exit Sub
                    Else
                        tEntrega.Text = lvVenta.SelectedItem.SubItems(5)
                        IngresoDeEntrega True
                        Foco tEntrega
                    End If
                End If
                
                
                BuscoCodigoEnCombo cCuota, ValorClave(lvVenta.SelectedItem.Key, "C")
                
                cArticulo.AddItem Trim(lvVenta.SelectedItem.SubItems(2))
                cArticulo.ItemData(cArticulo.NewIndex) = ValorClave(lvVenta.SelectedItem.Key, "A")
                cArticulo.ListIndex = 0
                cArticulo.Tag = ValorClave(lvVenta.SelectedItem.Key, "F")
                
                tCantidad.Text = lvVenta.SelectedItem.SubItems(1)
                tUnitario.Tag = lvVenta.SelectedItem.SubItems(3)        'P.U. Ctdo
                'Cuota
                If Not sConEntrega Then
                    tUnitario.Text = Format(CCur(lvVenta.SelectedItem.SubItems(6)) / CCur(tCantidad.Text), FormatoMonedaP)
                Else
                    tUnitario.Text = tUnitario.Tag
                End If
                
                Set oArtEdicion = CargoObjetoArticuloDeColeccion(ValorClave(lvVenta.SelectedItem.Key, "A"))
                
                If sConEntrega Then
'                    If Trim(lSubTotalF.Caption) = "" Then       'VALIDO QUE SE HAYA REALIZADO LA DISTRIBUCION
'                        MsgBox "Aún no ha ingresado el valor de entrega total de la solicitud o no se realizó la distribución automática.", vbExclamation, "ATENCIÓN"
'                        LimpioRenglon
'                        If tEntregaT.Enabled Then tEntregaT.SetFocus
'                    Else
'                        tEntrega.Text = lvVenta.SelectedItem.SubItems(5)
'                        IngresoDeEntrega True
'                        Foco tEntrega
'                    End If
                    
                Else
                    TotalesResto CCur(lvVenta.SelectedItem.SubItems(7)), CCur(lvVenta.SelectedItem.SubItems(4))
                    lvVenta.ListItems.Remove lvVenta.SelectedItem.Index
                End If
                Foco tCantidad
                AplicoCantidadLimitadaPorCantidad
                
                If lvVenta.ListItems.Count = 0 Then
                    cMoneda.Enabled = True
                    MnuEmitir.Enabled = False
                End If
                
            Case vbKeyDelete    'ELIMINO EL RENGLON----------------------------------------------------------------------------------
                'If Val(lArticulo.Tag) <> TFactura.Articulo Then Exit Sub
                If Val(lArticulo.Tag) <> TFactura.Articulo And Val(lArticulo.Tag) <> TFactura.ArtEspecifico Then Exit Sub
                
                If Trim(lvVenta.SelectedItem.SubItems(7)) <> "" Then
                    TotalesResto CCur(lvVenta.SelectedItem.SubItems(7)), CCur(lvVenta.SelectedItem.SubItems(4))
                End If
                
                'Si el que borro es plan con entrega REDISTRIBUYO------------------------------------
                If Left(lvVenta.SelectedItem.Key, 1) = "E" Then
                    lvVenta.ListItems.Remove lvVenta.SelectedItem.Index
                    If IsNumeric(tEntregaT.Text) And Trim(tEntregaT.Text) <> "" And lvVenta.ListItems.Count > 0 Then
                        DistribuirEntregas CCur(tEntregaT.Text)
                    Else
                        tEntregaT.Text = ""
                    End If
                    HabilitoEntrega
                Else
                    lvVenta.ListItems.Remove lvVenta.SelectedItem.Index
                    HabilitoEntrega
                End If
                
                If lvVenta.ListItems.Count > 0 Then
                    cMoneda.Enabled = False: MnuEmitir.Enabled = True
                Else
                    cMoneda.Enabled = True: MnuEmitir.Enabled = False
                End If
            
            Case vbKeyReturn: If tEntregaT.Enabled Then tEntregaT.SetFocus Else cPago.SetFocus
        End Select
    End If
    Exit Sub

ErrlvKD:
    clsGeneral.OcurrioError "Ocurrio un error inesperado."
End Sub

Private Sub MnuEmitir_Click()

    AccionGrabar

End Sub


Private Sub MnuHlp_Click()
On Error GoTo errHelp
    Screen.MousePointer = 11
    
    Dim aFile As String
    Cons = "Select * from Aplicacion Where AplNombre = '" & Trim(App.Title) & "'"
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

Private Sub MnuLimpiar_Click()
   LimpioTodaLaFicha True
End Sub

Private Sub LimpioTodaLaFicha(FocoCI As Boolean, Optional DejarGarantia As Boolean = False)
    On Error Resume Next
    LimpioRenglon
    
    Set colArtsGrilla = New Collection
    lvVenta.ListItems.Clear
    Set oArtEdicion = Nothing
    
    cMoneda.Enabled = True
    bSeñalInhabilitadoXMayor = False
    
    lblNombreCliente.Caption = ""
    labDireccion.Caption = "": labDireccion.Tag = ""
    lTEdad.Caption = ""
    
    cEMailsT.ClearObjects
    cEMailsT.Enabled = False
    cTelsT.Clear
    lTelsT.Caption = "Tels."
    cTelsT.BackColor = lblNombreCliente.BackColor
    cTelsT.ForeColor = lblNombreCliente.ForeColor
            
    lRucCliente.Caption = ""
    
        
    tUsuario.Text = "": tUsuario.Tag = ""
    tVendedor.Text = ""
    
    If Not DejarGarantia Then
        txtCliente.Text = ""
        txtGarantia.Text = ""
        lGarantia.Caption = "": lGEdad.Caption = ""
        gConyugeDelGarante = 0
    End If
    
    cCuota.Text = ""
    tEntregaT.Text = ""
    cComentario.Text = ""
    cPago.ListIndex = 0
    
    labTotal.Caption = "0.00": labIVA.Caption = "0.00": labSubTotal.Caption = "0.00"

    HabilitoEntrega
    IngresoDeEntrega False
            
    cDireccion.Clear: cDireccion.BackColor = labDireccion.BackColor
    gDirFactura = 0
    
    If FocoCI Then txtCliente.SetFocus
    
End Sub

Private Sub MnuPEmpleo_Click()
    If txtCliente.Cliente.Tipo = TC_Persona Then IrAEmpleo txtCliente.Cliente.Codigo
End Sub

Private Sub IrAEmpleo(ByVal Cliente As Long)
    
    On Error GoTo errIr
    Screen.MousePointer = 11
    Dim aObj As New clsCliente
    aObj.Empleos Cliente
    Set aObj = Nothing
    Screen.MousePointer = 0
    Exit Sub

errIr:
    clsGeneral.OcurrioError "Error al acceder al formulario de empleos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub IrAReferencia(Cliente As Long)
    On Error GoTo errIr
    Screen.MousePointer = 11
    Dim aObj As New clsCliente
    
    aObj.Referencias Cliente
    Set aObj = Nothing
    Screen.MousePointer = 0
    Exit Sub

errIr:
    clsGeneral.OcurrioError "Ocurrió un error al acceder al formulario.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub IrATitulo(Cliente As Long)
    On Error GoTo errIr
    Screen.MousePointer = 11
    Dim aObj As New clsCliente
    
    aObj.Titulos Cliente
    Set aObj = Nothing
    Screen.MousePointer = 0
    Exit Sub

errIr:
    clsGeneral.OcurrioError "Ocurrió un error al acceder al formulario.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuPReferencia_Click()
    If txtCliente.Cliente.Tipo = TC_Persona Then IrAReferencia txtCliente.Cliente.Codigo
End Sub

Private Sub MnuPTitulo_Click()
    If txtCliente.Cliente.Tipo = TC_Persona Then IrATitulo txtCliente.Cliente.Codigo
End Sub

Private Sub MnuVisulizacionOp_Click()
    On Error Resume Next
    EjecutarApp App.Path & "\Visualizacion de Operaciones", IIf(txtCliente.Cliente.Codigo > 0, CStr(txtCliente.Cliente.Codigo), "")
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Private Sub tCantidad_GotFocus()

    Foco tCantidad
    Status.Panels(1).Text = "Ingrese la cantidad de artículos."

End Sub

Private Sub tCantidad_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tCantidad.Text) Then
            If Val(lArticulo.Tag) = TFactura.ArtEspecifico Then
                If Val(tCantidad.Text) = 1 Then tUnitario.SetFocus
            ElseIf Val(tCantidad.Text) > 0 Then
                tUnitario.SetFocus
            End If
        End If
    End If

End Sub

Private Sub tCantidad_LostFocus()
    AplicoCantidadLimitadaPorCantidad
End Sub

Private Sub AplicoCantidadLimitadaPorCantidad()
    tmArticuloLimitado.Enabled = False
    If oArtEdicion.ID = 0 Then Exit Sub
    If IsNumeric(tCantidad.Text) Then
        If CInt(tCantidad.Text) > 0 Then
            tCantidad.Text = CInt(tCantidad.Text)
            tmArticuloLimitado.Enabled = (oArtEdicion.VentaXMayor = 0 And InStr(1, prmCategoriaDistribuidor, "," & Val(labDireccion.Tag) & ",") > 0) Or _
                oArtEdicion.VentaXMayor > 1 And oArtEdicion.VentaXMayor < Val(tCantidad.Text)
        Else
            tCantidad.Text = ""
        End If
    End If
    AplicoTextoDeVentaLimitada
End Sub

Private Sub tCantidad_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Status.Panels(1).Text = "Ingrese la cantidad de artículos."
End Sub

Private Sub tCantidad_Validate(Cancel As Boolean)
    
    If Not IsNumeric(tCantidad.Text) Then
        Cancel = True
        Foco tCantidad
    End If
    
    If Val(lArticulo.Tag) = TFactura.ArtEspecifico Then
        Cancel = (Val(tCantidad.Text) <> 1)
        If Cancel Then tCantidad.Text = 1: Foco tCantidad
    End If
End Sub

Private Sub LimpioDatosCliente()
On Error Resume Next
    
    lblNombreCliente.Caption = ""
    lTEdad.Caption = ""
    labDireccion.Caption = ""
    
    cEMailsT.ClearObjects
    cEMailsT.Enabled = False
    
    cTelsT.Clear
    lTelsT.Caption = "Tels."
    cTelsT.BackColor = lblNombreCliente.BackColor
    cTelsT.ForeColor = lblNombreCliente.ForeColor
    
    lRucCliente.Caption = ""
    
    cDireccion.Clear: cDireccion.BackColor = labDireccion.BackColor
    gDirFactura = 0
    
    txtGarantia.Text = ""
    lGarantia.Caption = ""
    lGEdad.Caption = ""
    
    labDireccion.Tag = ""
        
End Sub


Private Sub cComentario_GotFocus()
    cComentario.SelStart = 0: cComentario.SelLength = Len(cComentario.Text)
    Status.Panels(1).Text = "Ingrese un comentario para la solicitud."
End Sub

Private Sub cComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub tEntrega_GotFocus()
    tEntrega.SelStart = 0: tEntrega.SelLength = Len(tEntrega.Text)
    Status.Panels(1).Text = "Ingrese el valor de la entrega para el artículo seleccionado."
End Sub

Private Sub tEntrega_KeyPress(KeyAscii As Integer)

Dim aPlan As Long
Dim iAuxiliar As Currency

    If KeyAscii = vbKeyReturn And Trim(tEntrega.Text) <> "" Then
        If IsNumeric(tEntrega.Text) Then
            On Error GoTo errEntrega
            
            tEntrega.Text = Redondeo(tEntrega.Text, mMRound)
            
            'Valido que el importe de entrega sea menor a P.U. * Cantidad
            If (CCur(tUnitario.Text) * CCur(tCantidad.Text)) <= CCur(tEntrega.Text) Then
                MsgBox "El importe de entrega no debe superar a los precios contado de los artículos.", vbExclamation, "ATENCIÓN"
                Foco tEntrega
                Exit Sub
            End If  '----------------------------------------------------------------
            
            TotalesResto CCur(CCur(lvVenta.SelectedItem.SubItems(7))), CCur(lvVenta.SelectedItem.SubItems(4))
            
            tEntregaT.Text = Format(CCur(tEntregaT.Text) + (CCur(tEntrega.Text) - CCur(lvVenta.SelectedItem.SubItems(5))), "#,##0.00")
            lvVenta.SelectedItem.SubItems(5) = Format(tEntrega.Text, "#,##0.00")
            
            'El valor de la cuota es el (Precio Contado - Entrega) * Coeficiente ----- Coeficiente (Plan, TCuota, Moneda)
            aPlan = ValorClave(lvVenta.SelectedItem.Key, "P")
            
            Cons = "Select * from Coeficiente, TipoCuota" _
                & " Where CoePlan = " & aPlan _
                & " And CoeTipoCuota = " & cCuota.ItemData(cCuota.ListIndex) _
                & " And CoeMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                & " And CoeTipocuota = TCuCodigo"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            'Calculo lo que queda por pagar ((P.contado * Cantidad)- Entega * Coeficiente)
            iAuxiliar = ((CCur(tUnitario.Text) * CCur(tCantidad.Text)) - CCur(tEntrega.Text)) * RsAux!CoeCoeficiente
            
            'Veo si tiene descuento
            iAuxiliar = CCur(BuscoDescuentoCliente(cArticulo.ItemData(cArticulo.ListIndex), Val(labDireccion.Tag), iAuxiliar, _
                                                    CCur(tCantidad.Text), cArticulo.Text, cCuota.ItemData(cCuota.ListIndex)))
            
            'Valor de Cada Cuota
            lvVenta.SelectedItem.SubItems(6) = Format(Redondeo(iAuxiliar / RsAux!TCuCantidad, mMRound), "#,##0.00")
            'SubTotal = (Las cuotas + Entrega)
            lvVenta.SelectedItem.SubItems(7) = Format(CCur((lvVenta.SelectedItem.SubItems(6)) * RsAux!TCuCantidad) + CCur(tEntrega.Text), "#,##0.00")
            
            TotalesSumo CCur(lvVenta.SelectedItem.SubItems(7)), CCur(lvVenta.SelectedItem.SubItems(4))
            RsAux.Close
            IngresoDeEntrega False
            LimpioRenglon
        End If
    End If
    Exit Sub

errEntrega:
    clsGeneral.OcurrioError "Error al calcular los precios. Verifique los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tEntrega_LostFocus()

    IngresoDeEntrega False
    LimpioRenglon
            
End Sub

Private Sub tEntregaT_Change()
    sDistribuir = True
End Sub

Private Sub tEntregaT_GotFocus()

    sDistribuir = False
    tEntregaT.SelStart = 0
    tEntregaT.SelLength = Len(tEntregaT.Text)
    
    Status.Panels(1).Text = "Ingrese el valor total de la entrega para la solicitud."
    
End Sub

Private Sub tEntregaT_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn And Shift = vbCtrlMask Then
        If Trim(tEntregaT.Text) <> "" Then
            If IsNumeric(tEntregaT.Text) And sDistribuir Then
                tEntregaT.Text = Redondeo(CCur(tEntregaT.Text), mMRound)
                tEntregaT.Text = Format(tEntregaT.Text, "#,##0.00")
                
                DistribuirEntregas CCur(tEntregaT.Text)
                fnc_ValidoCtasPorEntrega True
                
            End If
            cPago.SetFocus
        End If
    End If
    
End Sub

Private Sub tEntregaT_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tEntregaT.Text) <> "" Then
            If IsNumeric(tEntregaT.Text) And sDistribuir Then
                tEntregaT.Text = Redondeo(CCur(tEntregaT.Text), mMRound)
                tEntregaT.Text = Format(tEntregaT.Text, "#,##0.00")
                DistribuirEntregas CCur(tEntregaT.Text)
                
                'Valido Valores Redondos de las Ctas    13/08/2004
                fnc_ValidoCtasPorEntrega False
            End If
            cPago.SetFocus
        End If
    End If
    
End Sub

Private Function fnc_ValidoCtasPorEntrega(ByVal preguntar As Boolean)
On Error GoTo errFncAjuste
    '---> Todos  los articulos deben tener el mismo Coef --> =Plan e = TipoCuota
    Dim itmP As ListItem
    Dim mIDPlan As Long, mIDTCuota As Long
    Dim mKeyP_TC As String, mVAux As String, bOK As Boolean
    mKeyP_TC = ""
    
    bOK = True
    For Each itmP In lvVenta.ListItems
        If mKeyP_TC = "" Then
            mKeyP_TC = CStr(ValorClave(itmP.Key, "P")) & CStr(ValorClave(itmP.Key, "C"))
            
            mIDPlan = ValorClave(itmP.Key, "P")
            mIDTCuota = ValorClave(itmP.Key, "C")
            
        Else
            mVAux = CStr(ValorClave(itmP.Key, "P")) & CStr(ValorClave(itmP.Key, "C"))
            If mVAux <> mKeyP_TC Then bOK = False
        End If
    Next
    
    '2) Voy a sumar todas las ctas para ver si da redondo
    If bOK Then
        Dim mUnaCta As Currency, mVCtas As Currency, mVCtasNEW As Currency, mVContados As Currency
        
        For Each itmP In lvVenta.ListItems
            mVContados = mVContados + (CCur(itmP.SubItems(3)) * CCur(itmP.SubItems(1)))
            mUnaCta = CCur(itmP.SubItems(6))
            mVCtas = mVCtas + mUnaCta
            
            If (mUnaCta Mod 10) <> 0 Then mUnaCta = Round(mUnaCta / 10, 0) * 10
            mVCtasNEW = mVCtasNEW + mUnaCta
        Next
        
        If mVCtasNEW <> mVCtas Then      'Posible Correccion
            Dim mNewEInicial As Currency: mNewEInicial = 0
            Dim mQCtas As Integer
            '1)  debo sacar la Qctas y coeficiente
            Cons = "Select CoeCoeficiente, TCuCantidad from Coeficiente, TipoCuota" _
                    & " Where CoePlan = " & mIDPlan _
                    & " And CoeTipoCuota = " & mIDTCuota _
                    & " And CoeMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                    & " And CoeTipocuota = TCuCodigo"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux.EOF Then
                '" Sum(PViPrecio)-(" & prmTotalFinanciado & "/CoeCoeficiente) as EInicial "
                mQCtas = RsAux!TCuCantidad
                mNewEInicial = mVContados - ((mVCtasNEW * mQCtas) / RsAux!CoeCoeficiente)
                mNewEInicial = Redondeo(mNewEInicial, mMRound)
            End If
            RsAux.Close
            
            If preguntar Then
                If MsgBox("Al entregar " & tEntregaT.Text & " quedan " & mQCtas & " cuotas de " & Format(mVCtas, "#,##0.00") & vbCrLf & vbCrLf & _
                               "Se sugiere hacer una entrega de " & Format(mNewEInicial, "#,##0.00") & "," & vbCrLf & _
                               "para que queden " & mQCtas & " cuotas de " & Format(mVCtasNEW, "#,##0.00") & _
                                vbCrLf & vbCrLf & _
                                "¿Desea aceptar la sugerencia?", vbQuestion + vbYesNo, "Cambiar Entrega") = vbNo Then
                                Exit Function
                End If
            Else
                MsgBox "Se modificó el valor de la entraga a $ " & Format(mNewEInicial, "#,##0.00") & " de forma que queden " & mQCtas & " cuotas de " & Format(mVCtasNEW, "#,##0.00"), vbInformation, "Entrega"
            End If
            tEntregaT.Text = Format(mNewEInicial, "#,##0.00")
            
            tEntregaT_KeyDown vbKeyReturn, vbCtrlMask
         End If
         
    End If
    Exit Function

errFncAjuste:
    clsGeneral.OcurrioError "Error al realizar el ajuste de las cuotas según el valor a entregar.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub tmArticuloLimitado_Timer()
    tmArticuloLimitado.Enabled = False
    If Val(tmArticuloLimitado.Tag) = 0 Then
        cArticulo.ForeColor = &HFF&
        tmArticuloLimitado.Tag = 1
    Else
        cArticulo.ForeColor = vbBlack
        tmArticuloLimitado.Tag = 0
    End If
    tCantidad.ForeColor = cArticulo.ForeColor
    tmArticuloLimitado.Enabled = True
End Sub

Private Sub tmClose_Timer()
On Error Resume Next
'    If DateDiff("s", CDate(tmClose.Tag), Now) > 14 Or oResAuto.Estado = 0 Then
        Unload Me
'    End If
End Sub


Private Function fnc_CargoDatosCliente(ByVal bConservarFactura As Boolean)
'ATENCION-------------------------------------------------------------------
'En el Tag de Dirección guardo la categoria de descuento del cliente.
'---------------------------------------------------------------------------
Dim rsCli As rdoResultset
Dim mSQL As String, bDataOK As Boolean
    
    mSQL = "Select CliCategoria, CPeApellido1, CPeApellido2, CPeNombre1, CPeNombre2, CPeFNacimiento, CEmFantasia, CEmNombre, CliDireccion " & _
            "From Cliente " & _
                " Left Outer Join CPersona On CliCodigo = CPeCliente" & _
                " Left Outer Join CEmpresa On CliCodigo = CEmCliente" & _
            " Where CliCodigo = " & txtCliente.Cliente.Codigo
            
    Set rsCli = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    bDataOK = Not rsCli.EOF
    
    Dim mTextoEmails As String
    If Not rsCli.EOF Then
        
        If Not bConservarFactura Or lvVenta.ListItems.Count = 0 Then
            LimpioTodaLaFicha False, True
        Else
            Dim idCatNew As Integer: idCatNew = 0
            If Not IsNull(rsCli!CliCategoria) Then idCatNew = rsCli!CliCategoria
            'Consulto si el cliente cambió de categoría.
            If CLng(labDireccion.Tag) <> idCatNew Then LimpioTodaLaFicha False, True
        End If

        If txtCliente.Cliente.Tipo = TC_Persona Then
            lRucCliente.Caption = txtCliente.Cliente.RutPersona
            lblNombreCliente.Caption = " " & ArmoNombre(Format(rsCli!CPeApellido1, "#"), Format(rsCli!CPeApellido2, "#"), Format(rsCli!CPeNombre1, "#"), Format(rsCli!CPeNombre2, "#"))
            mTextoEmails = Trim(rsCli!CPeApellido1) & ", " & Trim(rsCli!CPeNombre1)
            If Not IsNull(rsCli!CPeFNacimiento) Then lTEdad.Caption = CalculoEdad(rsCli("CPeFNacimiento")) '((Date - rsCli!CPeFNacimiento) \ 365)
        Else                                                                       'Tabla CEmrpesa -------------------------------------------------------------------
            lblNombreCliente.Caption = " " & Trim(rsCli!CEmFantasia)
            mTextoEmails = Trim(rsCli!CEmFantasia)
            If Not IsNull(rsCli!CEmNombre) Then lblNombreCliente.Caption = lblNombreCliente.Caption & " (" & Trim(rsCli!CEmNombre) & ")"
        End If
        
        If Not IsNull(rsCli!CliDireccion) Then
            cDireccion.AddItem "Dirección Principal": cDireccion.ItemData(cDireccion.NewIndex) = rsCli!CliDireccion
            cDireccion.Tag = rsCli!CliDireccion
            gDirFactura = rsCli!CliDireccion
        End If
    
        If Not IsNull(rsCli!CliCategoria) Then labDireccion.Tag = rsCli!CliCategoria
    Else
        LimpioTodaLaFicha False, True
    End If
    rsCli.Close
    
    If bDataOK Then
    
        CargoDireccionesAuxiliares txtCliente.Cliente.Codigo
        CargoCuotas IIf(Val(labDireccion.Tag) <> 0, Val(labDireccion.Tag), paCategoriaCliente)
        ListaDeSolicitudesPendientes txtCliente.Cliente.Codigo
        CargoTelefonos txtCliente.Cliente.Codigo
        'ValidoTelefonos fnd_IDCLiente
        cEMailsT.CargarDatos txtCliente.Cliente.Codigo
        cEMailsT.Enabled = True
        cEMailsT.IdsPorDefecto = StrConv(mTextoEmails, vbProperCase)
        
        On Error GoTo errVR
        Dim oValida As New clsValidaRUT
        If txtCliente.Cliente.Tipo = TC_Empresa And txtCliente.Cliente.Documento <> "" Then
            If Not oValida.ValidarRUT(txtCliente.Cliente.Documento) Then
                MsgBox "RUT INCORRECTO!!!, por favor valide con el cliente el número de RUT ya que no cumple con la validación.", vbExclamation, "RUT INCORRECTO"
            End If
        ElseIf txtCliente.Cliente.Tipo = TC_Persona And txtCliente.Cliente.RutPersona <> "" Then
            If Not oValida.ValidarRUT(txtCliente.Cliente.RutPersona) Then
                MsgBox "RUT INCORRECTO!!!, por favor valide con el cliente el número de RUT ya que no cumple con la validación.", vbExclamation, "RUT INCORRECTO"
            End If
        End If
        Set oValida = Nothing
        
    End If
    Exit Function
errVR:
    clsGeneral.OcurrioError "Error al validar el RUT", Err.Description, "Validar RUT"
End Function

Private Sub fnc_CargoDatosGarantia()
Dim rsCli As rdoResultset
Dim mSQL As String
    
    lGarantia.Caption = "": lGEdad.Caption = ""
    mSQL = "Select CPeApellido1, CPeApellido2, CPeNombre1, CPeNombre2, CPeFNacimiento, CEmFantasia, CEmNombre, CPeConyuge From Cliente " & _
                " Left Outer Join CPersona On CliCodigo = CPeCliente" & _
                " Left Outer Join CEmpresa On CliCodigo = CEmCliente" & _
            " Where CliCodigo = " & txtGarantia.Cliente.Codigo
    Set rsCli = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    If Not rsCli.EOF Then
        
        If txtGarantia.Cliente.Tipo = TC_Persona Then
            lGarantia.Caption = " " & ArmoNombre(Format(rsCli!CPeApellido1, "#"), Format(rsCli!CPeApellido2, "#"), Format(rsCli!CPeNombre1, "#"), Format(rsCli!CPeNombre2, "#"))
            lGEdad.Tag = 1
        Else
            lGarantia.Caption = " " & Trim(rsCli!CEmFantasia)
            lGEdad.Tag = 0
        End If
        If Not IsNull(rsCli!CPeFNacimiento) Then lGEdad.Caption = CalculoEdad(rsCli!CPeFNacimiento)
        If Not IsNull(rsCli!CPeConyuge) Then gConyugeDelGarante = rsCli!CPeConyuge Else gConyugeDelGarante = 0
    End If
    rsCli.Close
    
End Sub

Private Sub BuscoArticuloXNombre()
On Error GoTo ErrBAN

    Screen.MousePointer = 11
    Dim aIdSeleccionado As Long, aArticulo As Long, aQ As Integer
    'cArticulo.Text = clsGeneral.Replace(cArticulo.Text, " ", "%")
    aQ = 0: aIdSeleccionado = 0
    
    Cons = "SELECT dbo.ocxarticulo('Solicitud', '" & cCuota.ItemData(cCuota.ListIndex) & "<QC>" & cArticulo.Text & "')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        Cons = RsAux(0)
    Else
        Cons = ""
    End If
    RsAux.Close
    If Cons <> "" Then
        Dim objLista As New clsListadeAyuda
        objLista.CerrarSiEsUnico = True
        aIdSeleccionado = objLista.ActivarAyuda(cBase, Cons, 5200, 1)
    End If
    Me.Refresh
    If aIdSeleccionado > 0 Then
        aIdSeleccionado = objLista.RetornoDatoSeleccionado(0)
    End If
    Set objLista = Nothing
    
    cArticulo.Clear
    If aIdSeleccionado > 0 Then
        BuscoArticuloxCodigo 0, aIdSeleccionado
    Else
        tUnitario.Text = ""
    End If
    Screen.MousePointer = 0
    Exit Sub
    
ErrBAN:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub BuscoArticuloEspecifico(Texto As String)
'   lSubTotalF.Tag  = Precio Unitario Financiado
'   cCuota.Tag       = Cantidad de Cuotas
'   tCantidad.Tag   = Plan
'   tUnitario.tag = Precio Unitario Contado        ------ y el en Text va el valor de la cuota

On Error GoTo ErrBAC

    Screen.MousePointer = 11
    cArticulo.Clear
    tUnitario.Text = "": tUnitario.Tag = ""
    lSubTotalF.Tag = ""
    
    'Saco los datos del Articulo--------------------------------------------------------------------------
    Dim pstrSQL As String, pbytQ As Byte
    Dim plngArticulo As Long, plngArtEspecifico As Long, pstrNombre As String, pcurVariacionPrecio As Currency
    
'    pstrSQL = "SELECT ArtId, AEsID as ID, AEsNombre as Nombre, AEsVariacionPrecio as 'Variación Precio'" & _
'            " FROM ArticuloEspecifico inner join Articulo on AesArticulo = ArtId" & _
'            " WHERE (AEsNombre like '[p1]%' OR ArtNombre like '[p1]%'" & _
'                    " OR Convert(varchar(10), ArtCodigo) = '[p1]'" & _
'                    " OR Convert(varchar(10), AEsId) = '[p1]' )" & _
'            " And AEsEstado = 1 " & _
'            " And AEsDocumento IS NULL"
'
'    pstrSQL = Replace(pstrSQL, "[p1]", Texto)
'
'    Set RsAux = cBase.OpenResultset(pstrSQL, rdOpenDynamic, rdConcurValues)
'    If Not RsAux.EOF Then
'        plngArticulo = RsAux!ArtID
'        plngArtEspecifico = RsAux(1)    'AEsID
'        pstrNombre = Trim(RsAux(2))     'AEsNombre
'        pcurVariacionPrecio = RsAux(3)  'AEsVariacionPrecio
'        pbytQ = 1
'
'        RsAux.MoveNext
'        If Not RsAux.EOF Then pbytQ = 2
'    End If
'    RsAux.Close
'
'    Select Case pbytQ
'        Case 0: MsgBox "No existen artículos específicos para los datos ingresados.", vbInformation, "No hay Datos"
'
'        Case 2:
'            Dim objLista As New clsListadeAyuda
'            plngArticulo = objLista.ActivarAyuda(cBase, pstrSQL, 5200, 1)
'            Me.Refresh
'            If plngArticulo > 0 Then
'                plngArticulo = objLista.RetornoDatoSeleccionado(0)
'                plngArtEspecifico = objLista.RetornoDatoSeleccionado(1)
'                pstrNombre = Trim(objLista.RetornoDatoSeleccionado(2))
'                pcurVariacionPrecio = objLista.RetornoDatoSeleccionado(3)
'            End If
'            Set objLista = Nothing
'    End Select

    Dim objLista As New clsListadeAyuda
    plngArticulo = objLista.ActivarAyuda(cBase, "EXEC prg_BuscarArticuloEspecifico '" + Texto + "'", 5200, 3)
    Me.Refresh
    If plngArticulo > 0 Then
        plngArticulo = objLista.RetornoDatoSeleccionado(0)
        plngArtEspecifico = objLista.RetornoDatoSeleccionado(3)
        pstrNombre = Trim(objLista.RetornoDatoSeleccionado(4))
        pcurVariacionPrecio = objLista.RetornoDatoSeleccionado(5)
    End If
    Set objLista = Nothing

    Screen.MousePointer = 0
    
    If plngArticulo = 0 Then Exit Sub
    cArticulo.Clear
               
'    'Verifico si el articulo está en la lista---------------------------------------------------------------------------------
    If Ingresado(plngArticulo, cCuota.ItemData(cCuota.ListIndex)) Then
        Screen.MousePointer = 0
        MsgBox "El artículo seleccionado ya fue ingresado. Para modificarlo edítelo.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    '--------------------------------------------------------------------------------------------------------------------------
    cArticulo.AddItem pstrNombre
    cArticulo.ItemData(cArticulo.NewIndex) = plngArticulo
    cArticulo.ListIndex = 0
    cArticulo.Tag = plngArtEspecifico
    
    Set oArtEdicion = New clsArticulo
    oArtEdicion.NombreArticulo = pstrNombre
    oArtEdicion.VentaXMayor = 1
    oArtEdicion.ID = plngArticulo
    
    
    '------------------------------------------------------------------------------------------------------------------
    
    Dim miUnitarioF As Currency, miCuotaF As Currency, bNoHabPlan As Boolean, miPlan As Long, bOK As Boolean
    bOK = PrecioArticulo(plngArticulo, sConEntrega, cMoneda.ItemData(cMoneda.ListIndex), cCuota.ItemData(cCuota.ListIndex), Val(cCuota.Tag), _
                         miUnitarioF, miCuotaF, miPlan, bNoHabPlan, pcurVariacionPrecio)
    
    If Not bOK Then
        MsgBox "No existe un coeficiente para el cálculo de cuotas. " & vbCrLf & _
               "El artículo no se puede vender en esta financiación, consulte.", vbExclamation, "Falta Coeficiente"
        cArticulo.Clear
        tUnitario.Text = "": tUnitario.Tag = ""
    
    Else
        'miUnitarioF = miUnitarioF + pcurVariacionPrecio
        If Not sConEntrega Then         'PROCESO PLAN SIN ENTREGA------------------------------------------------
         '   miCuotaF = miUnitarioF / Val(cCuota.Tag)
            If bNoHabPlan Then
                If MsgBox("El artículo no está habilitado para la venta en '" & Trim(cCuota.Text) & "'" & vbCrLf & _
                          "Desea hacerlo igualmente ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Artículo No Habilitado en el Plan") = vbNo Then
                    cArticulo.Clear
                End If
            End If
            
            If cArticulo.ListCount > 0 Then
                lSubTotalF.Tag = miUnitarioF
                tUnitario.Text = Format(miCuotaF, "#,##0.00")     'Valor Cuota Finanaciado
                tCantidad.Tag = miPlan
            End If
            
        Else                                        'PROCESO PLAN CON  ENTREGA------------------------------------------------
            tUnitario.Tag = miUnitarioF                                     'Precio Unitario Contado
            tUnitario.Text = Format(miUnitarioF, "#,##0.00")    'Precio Unitario Contado
            tCantidad.Tag = miPlan
        End If
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
ErrBAC:
    clsGeneral.OcurrioError "Error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub BuscoArticuloxCodigo(Codigo As Long, ID As Long)

'   lSubTotalF.Tag  = Precio Unitario Financiado
'   cCuota.Tag       = Cantidad de Cuotas
'   tCantidad.Tag   = Plan
'   tUnitario.tag = Precio Unitario Contado        ------ y el en Text va el valor de la cuota

On Error GoTo ErrBAC
Dim bEsCombo As Boolean

    bEsCombo = False
    Screen.MousePointer = 11
    Set oArtEdicion = New clsArticulo
    
    cArticulo.Clear
    tUnitario.Text = "": tUnitario.Tag = ""
    lSubTotalF.Tag = ""
    
    
    'Saco los datos del Articulo--------------------------------------------------------------------------
    If ID > 0 Then
        Cons = "Select * From Articulo Where ArtID = " & ID
    Else
        Cons = "Select * From Articulo Where ArtCodigo = " & Codigo
    End If
    
    Set RsArt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not RsArt.EOF Then
        '1)      Si el articulo es del Tipo Flete lo cambio por el del parametro prmArticuloFleteVenta
        If fnc_EsDelTipoFlete(RsArt!ArtID) Then
            RsArt.Close
            Cons = "Select * From Articulo Where ArtID = " & prmArticuloFleteVenta
            Set RsArt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        End If
    End If
    
    If Not RsArt.EOF Then
        If RsArt!ArtEsCombo Then bEsCombo = True
        
        If RsArt!ArtEnUso = 0 Then
            'MsgBox "Artículo: " & Format(RsArt!ArtCodigo, "(#,000,000)") & " " & Trim(RsArt!ArtNombre) & Chr(vbKeyReturn) & "El artículo seleccionado no está en uso. Consulte.", vbExclamation, "ATENCIÓN"
            'RsArt.Close: Screen.MousePointer = 0: Exit Sub
            snd_ActivarSonido prmPathSound & sndArtFueraUso
            If MsgBox("Artículo: " & Format(RsArt!ArtCodigo, "(#,000,000)") & " " & Trim(RsArt!ArtNombre) & vbCrLf & _
                           "El artículo seleccionado no está en uso. Consulte." & vbCrLf & vbCrLf & _
                           "Para agregarlo a la solicitud presione SI.", vbExclamation + vbYesNo + vbDefaultButton2, "Artículo Fuera de Uso") = vbNo Then
                RsArt.Close: Screen.MousePointer = 0: Exit Sub
            End If
            
        Else
            If IsNull(RsArt!ArtHabilitado) Or UCase(RsArt!ArtHabilitado) <> "S" Then
                snd_ActivarSonido prmPathSound & sndArtNoHabilitado
                If MsgBox("Artículo: " & Format(RsArt!ArtCodigo, "(#,000,000)") & " " & Trim(RsArt!ArtNombre) & vbCrLf & _
                               "El artículo no está habilitado para la venta." & vbCrLf & vbCrLf & _
                               "Para agregarlo a la solicitud presione SI.", vbExclamation + vbYesNo + vbDefaultButton2, "Artículo NO Habilitado") = vbNo Then
                    RsArt.Close: Screen.MousePointer = 0: Exit Sub
                End If
            End If
        End If
            
    Else
        MsgBox "El artículo ingresado no existe. Verifique los datos.", vbExclamation, "ATENCIÓN"
        RsArt.Close: Screen.MousePointer = 0: Exit Sub
    End If
    
    'Verifico si el articulo está en la lista---------------------------------------------------------------------------------
    If Ingresado(RsArt!ArtID, cCuota.ItemData(cCuota.ListIndex)) Then
        Screen.MousePointer = 0
        MsgBox "El artículo seleccionado ya fue ingresado. Para modificarlo edítelo.", vbExclamation, "ATENCIÓN"
        RsArt.Close: Exit Sub
    End If
    '--------------------------------------------------------------------------------------------------------------------------
    
    Set oArtEdicion = New clsArticulo
    oArtEdicion.ID = RsArt!ArtID
    If Not IsNull(RsArt("ArtEnVentaXMayor")) Then
                'Guardo la cantidad de arts posibles a vender al por mayor.
        oArtEdicion.VentaXMayor = RsArt("ArtEnVentaXMayor")
    Else
        oArtEdicion.VentaXMayor = 1
    End If
    oArtEdicion.NombreArticulo = Trim(RsArt!ArtNombre)
    
    cArticulo.AddItem Trim(RsArt!ArtNombre)
    cArticulo.ItemData(cArticulo.NewIndex) = RsArt!ArtID
    cArticulo.ListIndex = 0
    Codigo = RsArt!ArtID
    RsArt.Close
    '------------------------------------------------------------------------------------------------------------------
    
    AplicoTextoDeVentaLimitada
    
    If bEsCombo Then
        ProcesoArticuloCombo Codigo
        Screen.MousePointer = 0: Exit Sub
    End If
    
    Dim miUnitarioF As Currency, miCuotaF As Currency, bNoHabPlan As Boolean, miPlan As Long, bOK As Boolean
    bOK = PrecioArticulo(Codigo, sConEntrega, cMoneda.ItemData(cMoneda.ListIndex), cCuota.ItemData(cCuota.ListIndex), Val(cCuota.Tag), _
                        miUnitarioF, miCuotaF, miPlan, bNoHabPlan)
    
    If Not bOK Then
        Set oArtEdicion = New clsArticulo
        MsgBox "No existe un coeficiente para el cálculo de cuotas. " & vbCrLf & "El artículo no se puede vender en esta financiación, consulte.", vbExclamation, "Falta Coeficiente"
        cArticulo.Clear
        tUnitario.Text = "": tUnitario.Tag = ""
    
    Else
        
        If Not sConEntrega Then         'PROCESO PLAN SIN ENTREGA------------------------------------------------
            If bNoHabPlan Then
                If MsgBox("El artículo no está habilitado para la venta en '" & Trim(cCuota.Text) & "'" & vbCrLf & _
                          "Desea hacerlo igualmente ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Artículo No Habilitado en el Plan") = vbNo Then
                    cArticulo.Clear
                End If
            End If
            
            If cArticulo.ListCount > 0 Then
                lSubTotalF.Tag = miUnitarioF 'Format(miUnitarioF, "#,##0")    'Precio de Unitario Financiaciado
                tUnitario.Text = Format(miCuotaF, "#,##0.00")     'Valor Cuota Finanaciado
                tCantidad.Tag = miPlan
            End If
            
        Else                                        'PROCESO PLAN CON  ENTREGA------------------------------------------------
            tUnitario.Tag = miUnitarioF                                     'Precio Unitario Contado
            tUnitario.Text = Format(miUnitarioF, "#,##0.00")    'Precio Unitario Contado
            tCantidad.Tag = miPlan
        End If
            
'        If cArticulo.ListIndex >= 0 Then
'
'            If txtCliente.Cliente.Codigo > 0 And InStr(1, prmCategoriaDistribuidor, "," & Val(labDireccion.Tag) & ",") > 0 And oArtEdicion.VentaXMayor = 0 Then
'                If MsgBox("El artículo ingresado no está habilitado para vender al por mayor." & vbCrLf & "¿Desea facturarlo de todas formas?", vbQuestion + vbYesNo + vbDefaultButton2, "Ventas por mayor") = vbNo Then
'                    RsAux.Close
'                    Set oArtEdicion = Nothing
'                    cArticulo.Clear
'                End If
'            End If
'
'        End If
        
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
ErrBAC:
    clsGeneral.OcurrioError "Error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AplicoTextoDeVentaLimitada()
    
    If IsNumeric(tCantidad.Text) Then
    
        If oArtEdicion.VentaXMayor = 0 And (InStr(1, prmCategoriaDistribuidor, "," & Val(labDireccion.Tag) & ",") > 0 Or Val(labDireccion.Tag) = 0) Then
            cArticulo.List(0) = oArtEdicion.NombreArticulo & " (no vta. a Distr.)"
        ElseIf oArtEdicion.VentaXMayor > 1 And oArtEdicion.VentaXMayor < Val(tCantidad.Text) Then
            cArticulo.List(0) = oArtEdicion.NombreArticulo & " (limitado a " & oArtEdicion.VentaXMayor & ")"
        Else
            cArticulo.List(0) = oArtEdicion.NombreArticulo
        End If
        
    Else
        
        'Defino el nombre en base a la disponibilidad de venta.
        If oArtEdicion.VentaXMayor = 0 And (InStr(1, prmCategoriaDistribuidor, "," & Val(labDireccion.Tag) & ",") > 0 Or Val(labDireccion.Tag) = 0) Then
            cArticulo.List(0) = oArtEdicion.NombreArticulo & " (no vta. a Distr.)"
        ElseIf oArtEdicion.VentaXMayor > 1 Then
            cArticulo.List(0) = oArtEdicion.NombreArticulo & " (limitado a " & oArtEdicion.VentaXMayor & ")"
        Else
            cArticulo.List(0) = oArtEdicion.NombreArticulo
        End If
    End If
    
'    If Not tmArticuloLimitado.Enabled Then
        tmArticuloLimitado.Enabled = (cArticulo.List(0) <> oArtEdicion.NombreArticulo)
        If Not tmArticuloLimitado.Enabled Then
            cArticulo.ForeColor = vbBlack
            tCantidad.ForeColor = vbBlack
        End If
'    End If
    
End Sub


Private Sub ProcesoArticuloCombo(IDArticulo As Long, Optional bInsertarFilas As Boolean = False, Optional ImporteDigitado As Currency = 0)
On Error GoTo errPCombo

Dim arrCombo() As typCombo

Dim cbTotalF As Currency    'Total financiado del combo
Dim cbValorCuotaF As Currency   'Valor de c/cta del combo

Dim arTotalF As Currency    'Total financiado del articulo
Dim arValorCuotaF As Currency   'Valor de c/cta del articulo

Dim RsCom As rdoResultset
Dim iCount As Integer, miIDArticulo As Long, miQArticulo As Integer
    
    Dim aCostoB As Currency
    ReDim arrCombo(0)
    iCount = 1
    Cons = "Select * from Presupuesto, PresupuestoArticulo  " & _
               " Where PreArtCombo = " & IDArticulo & _
               " And PreID = PArPresupuesto" & _
               " And PreMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
    Set RsCom = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not RsCom.EOF Then
        miIDArticulo = RsCom!PreArticulo
        aCostoB = RsCom!PreImporte
        
        Do While Not RsCom.EOF
            iCount = UBound(arrCombo) + 1
            ReDim Preserve arrCombo(iCount)
            arrCombo(iCount).Articulo = RsCom!PArArticulo
            arrCombo(iCount).Q = RsCom!PArCantidad
            arrCombo(iCount).EsBonificacion = False
            arrCombo(iCount).Bonificacion = 0
            RsCom.MoveNext
        Loop
        
        iCount = UBound(arrCombo) + 1
        ReDim Preserve arrCombo(iCount)
        arrCombo(iCount).Articulo = miIDArticulo
        arrCombo(iCount).Q = 1
        arrCombo(iCount).EsBonificacion = True
        arrCombo(iCount).Bonificacion = aCostoB
                
    End If
    RsCom.Close
    
    'PreArticulo es el articulo Bonificacion
    
    Dim miUnitarioF As Currency, miCuotaF As Currency, bNoHabPlan As Boolean, miPlan As Long, bOK As Boolean
    Dim aCoefPD  As Currency
    
    For i = 1 To UBound(arrCombo)
        arTotalF = 0: arValorCuotaF = 0
                
        miIDArticulo = arrCombo(i).Articulo
        miQArticulo = arrCombo(i).Q
            
        If Not sConEntrega Then         'PROCESO PLAN SIN ENTREGA------------------------------------------------
            
            If arrCombo(i).EsBonificacion Then   'Proceso el articulo bonificacion
                aCoefPD = 1
                arTotalF = arrCombo(i).Bonificacion
                If arTotalF <> 0 Then
                    'Busco el coeficiente p/tipo de Cuota y plan por defecto para finaciar la bonificacion
                    Cons = "Select * from Coeficiente" & _
                                " Where CoePlan = " & paPlanPorDefecto & _
                                " And CoeTipoCuota = " & cCuota.ItemData(cCuota.ListIndex) & _
                                " And CoeMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    If Not RsAux.EOF Then aCoefPD = RsAux!CoeCoeficiente
                    RsAux.Close
                End If
                
                'arTotalF = Format(arTotalF * aCoefPD, "0")
                arTotalF = Redondeo(arTotalF * aCoefPD, mMRound)
                arTotalF = arTotalF * miQArticulo
                arTotalF = Redondeo(arTotalF / cCuota.Tag, mMRound)
                arTotalF = arTotalF * cCuota.Tag
                arValorCuotaF = Format(arTotalF / cCuota.Tag, "#,##0.00")        'Valor Cuota Finanaciado
                
                If cbTotalF + arTotalF <> ImporteDigitado And bInsertarFilas Then
                    arTotalF = arTotalF + ImporteDigitado - (cbTotalF + arTotalF)
                    arValorCuotaF = Format(arTotalF / cCuota.Tag, "#,##0.00")
                End If
                
            Else
                
                bOK = PrecioArticulo(miIDArticulo, sConEntrega, cMoneda.ItemData(cMoneda.ListIndex), cCuota.ItemData(cCuota.ListIndex), Val(cCuota.Tag), _
                            miUnitarioF, miCuotaF, miPlan, bNoHabPlan)
                
                If Not bOK Then Exit For
                
                arTotalF = miUnitarioF                          'Precio de Unitario Financiaciado
                arTotalF = arTotalF * miQArticulo
                arValorCuotaF = Format(miCuotaF * miQArticulo, "#,##0.00")     'Valor Cuota Finanaciado
                tCantidad.Tag = miPlan
            End If
            
        Else                                            'PROCESO PLAN CON ENTREGA------------------------------------------------
            
            If arrCombo(i).EsBonificacion Then   'Proceso el articulo bonificacion
                arTotalF = arrCombo(i).Bonificacion
                If cbTotalF + arTotalF <> ImporteDigitado And bInsertarFilas Then
                    arTotalF = arTotalF + ImporteDigitado - (cbTotalF + arTotalF)
                End If
            
            Else
                
                bOK = PrecioArticulo(miIDArticulo, sConEntrega, cMoneda.ItemData(cMoneda.ListIndex), cCuota.ItemData(cCuota.ListIndex), Val(cCuota.Tag), _
                                                miUnitarioF, miCuotaF, miPlan, bNoHabPlan)

                If Not bOK Then Exit For
                
                arTotalF = miUnitarioF    'Precio Unitario Contado
                arTotalF = arTotalF * miQArticulo
                tCantidad.Tag = miPlan
                
            End If
        End If
        
        If Not bOK Then Exit For
        
        If bInsertarFilas And Not arrCombo(i).EsBonificacion And Not sConEntrega Then
            arTotalF = arTotalF * CCur(tCantidad.Text)
            arValorCuotaF = arValorCuotaF * CCur(tCantidad.Text)
        End If
        cbTotalF = cbTotalF + arTotalF
        cbValorCuotaF = cbValorCuotaF + arValorCuotaF
        
        'Si es p/Insertar-----------------------------------------------------------------------------------------------------------------------------------------------------
        If bInsertarFilas Then
            If (arrCombo(i).EsBonificacion And arTotalF <> 0) Or (Not arrCombo(i).EsBonificacion) Then
                'Si la bon es cero y no me cambio los costos ---> no iserto
                If sConEntrega Then
                    Set itmX = lvVenta.ListItems.Add(, "E" & cCuota.ItemData(cCuota.ListIndex) & "P" & tCantidad.Tag & "A" & miIDArticulo, cCuota.Text)
                    itmX.SubItems(8) = arTotalF / miQArticulo  'Unitario Contado (para controlar cambio de precios con entrega)
                    itmX.SubItems(3) = Format(arTotalF / miQArticulo, "#,##0.00")                     'Contado
                Else
                    Set itmX = lvVenta.ListItems.Add(, "N" & cCuota.ItemData(cCuota.ListIndex) & "P" & tCantidad.Tag & "A" & miIDArticulo, cCuota.Text)
                    itmX.SubItems(3) = Format("", "#,##0.00")                     'Contado
                End If
                
                itmX.SubItems(1) = miQArticulo * CCur(tCantidad.Text)
                
                Cons = "Select * from Articulo Where ArtId = " & miIDArticulo
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsAux.EOF Then
                    itmX.SubItems(2) = Trim(RsAux!ArtNombre)
                End If
                RsAux.Close
                itmX.SubItems(4) = IVAArticulo(miIDArticulo)
                
                If Trim(lSubTotalF.Caption) <> "" Then
                    itmX.SubItems(6) = Format(arValorCuotaF, FormatoMonedaP)        'Cuota
                    
                    'Ajusto el subtotal con lo que me da la cuota (SubTotal)
                    itmX.SubItems(7) = Format(arTotalF, FormatoMonedaP)    'Total Financiado
                    
                    itmX.SubItems(8) = arTotalF / CCur(tCantidad.Text) / miQArticulo 'Unitario Financiado
                    
                    TotalesSumo CCur(itmX.SubItems(7)), CCur(itmX.SubItems(4))
                End If
            End If
        End If
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Next
    
    If Not bOK Then
        MsgBox "No existe un coeficiente para el cálculo de cuotas. " & vbCrLf & _
                    "El artículo no se puede vender en esta financiación, consulte.", vbExclamation, "Falta Coeficiente"
        cArticulo.Clear
        tUnitario.Text = "": tUnitario.Tag = ""
    Else
    
        If Not bInsertarFilas Then
            If Not sConEntrega Then
                lSubTotalF.Tag = cbTotalF 'Format(cbTotalF, "#,##0")                          'Precio de Unitario Financiaciado
                tUnitario.Text = Format(cbTotalF / cCuota.Tag, "#,##0.00")     'Valor Cuota Finanaciado
            Else
                tUnitario.Tag = cbTotalF    'Precio Unitario Contado
                tUnitario.Text = Format(cbTotalF, FormatoMonedaP)   'Precio Unitario Contado
            End If
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errPCombo:
    clsGeneral.OcurrioError "Error al procesar los artículos del combo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tUnitario_Change()
    lSubTotalF.Caption = ""
End Sub

Private Sub tUnitario_GotFocus()
On Error Resume Next

Dim aUnitarioF As Currency
    
    Status.Panels(1).Text = "Precio contado unitario del artículo."
    tUnitario.SelStart = 0: tUnitario.SelLength = Len(tUnitario.Text)
    
    If sConEntrega Or Trim(tCantidad.Text) = "" Or Not IsNumeric(tCantidad.Text) Or cArticulo.ListIndex = -1 Then Exit Sub
    
    If lSubTotalF.Tag <> "" Then      'En el TAG Tengo el Precio financiado del articulo
        
        aUnitarioF = BuscoDescuentoCliente(cArticulo.ItemData(cArticulo.ListIndex), Val(labDireccion.Tag), CCur(lSubTotalF.Tag), _
                            Val(tCantidad.Text), cArticulo.Text, cCuota.ItemData(cCuota.ListIndex))
        tUnitario.Text = Redondeo(aUnitarioF / CCur(cCuota.Tag), mMRound)
        tUnitario.Text = Format(tUnitario.Text, FormatoMonedaP)
        lSubTotalF.Caption = Format(((tUnitario.Text * CCur(cCuota.Tag)) * CCur(tCantidad.Text)), FormatoMonedaP)
        tUnitario.SelStart = 0: tUnitario.SelLength = Len(tUnitario.Text)
    End If

End Sub

Private Function ObtenerCoeficiente(ByVal plan As Integer, ByVal tipocuota As Integer, ByVal Moneda As Integer) As Currency
Dim rsC As rdoResultset
    ObtenerCoeficiente = 0
    Cons = "Select CoeCoeficiente from Coeficiente" _
            & " Where CoePlan = " & plan _
            & " And CoeTipoCuota = " & tipocuota _
            & " And CoeMoneda = " & Moneda
    Set rsC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsC.EOF Then    'Si NO hay coeficientes NO SE VENDE
        ObtenerCoeficiente = rsC(0)
    End If
    rsC.Close
End Function

Private Sub tUnitario_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF1 And cArticulo.ListIndex > -1 And Val(tUnitario.Tag) = 0 And sConEntrega Then
        On Error GoTo errKD
        Dim sImp As String: sImp = InputBox("Ingrese el monto de la entrega", "Calcular contado", "")
        Dim iEnt As Currency, iCuota As Currency
        If IsNumeric(sImp) Then
            iEnt = CCur(sImp)
            sImp = InputBox("Ingrese el valor de la cuota", "Calcular contado", "")
            If IsNumeric(sImp) Then
                iCuota = CCur(sImp)

                Dim iCoef As Currency: iCoef = ObtenerCoeficiente(Val(tCantidad.Tag), cCuota.ItemData(cCuota.ListIndex), cMoneda.ItemData(cMoneda.ListIndex))

                'Calculo del contado.
                'Carlos Gutierrez dice:
                'Cdo = (Cta * QC) / Coef + Ent
                
                If iCoef > 0 Then
                    iEnt = ((iCuota * Val(cCuota.Tag)) / iCoef) + iEnt
                    tUnitario.Text = Format(iEnt, "#,##.00")
                End If
            End If
        End If
    End If
    Exit Sub
errKD:
    clsGeneral.OcurrioError "Error al calcular.", Err.Description
End Sub

Private Sub tUnitario_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And IsNumeric(tUnitario.Text) And cMoneda.ListIndex > -1 Then
        If Not IsNumeric(tCantidad.Text) Then tCantidad.Text = 1
        If IsNumeric(tUnitario.Text) Then
            If Not sConEntrega Then
                'Hay que calcular los totales -- por si cambió la cuota.
                tUnitario.Text = Redondeo(CCur(tUnitario.Text), mMRound)
                lSubTotalF.Caption = Format((CCur(tUnitario.Text) * CCur(cCuota.Tag)) * CCur(tCantidad.Text), FormatoMonedaP)
            Else
                'tUnitario.Tag = tUnitario.Text  'Por si modifico el unitario ctdo.
            End If
            InsertoFila
        End If
    Else
        If cMoneda.ListIndex = -1 Then
            MsgBox "No seleccionó una moneda.", vbCritical, "ATENCIÓN"
            LimpioRenglon
            cMoneda.SetFocus
            lvVenta.ListItems.Clear
            Set colArtsGrilla = New Collection
        End If
    End If
        
End Sub

Private Sub LimpioRenglon()

    cArticulo.Clear
    tCantidad.Text = ""
    tUnitario.Text = ""
    tEntrega.Text = ""
    lSubTotalF.Caption = ""

    tUnitario.Tag = ""
    lSubTotalF.Tag = ""
    tEntrega.Tag = ""

    Set oArtEdicion = New clsArticulo
    tmArticuloLimitado.Enabled = False
    cArticulo.ForeColor = vbBlack
    tCantidad.ForeColor = vbBlack

End Sub

'--------------------------------------------------------------------------------------------------------------------------------------------------
'   Parametros:
'       CodArticulo = id de Articulo.
'       CodCatCliente = categoria de cliente
'       Unitario = Precio del articulo para aplicar el dto.
'       Cantidad = cantidad de articulos para validar la minima.
'       Articulo = texto del articulo para dar mensaje.
'       Plan = id de plan
'
'   RETORNA = precio unitario con o sin descuentos
'--------------------------------------------------------------------------------------------------------------------------------------------------
Private Function BuscoDescuentoCliente(CodArticulo As Long, CodCatCliente As Long, Unitario As Currency, Cantidad As Currency, _
                                                          Articulo As String, plan As Long) As String

On Error GoTo errDescuento

Dim RsBDC As rdoResultset
Dim aRetorno As Currency
    
    
    aRetorno = Redondeo(Unitario, mMRound)
    
    If CodCatCliente > 0 Then
    
        Cons = "Select CDTPorcentaje, AFaCantidadD From ArticuloFacturacion, CategoriaDescuento" _
                & " Where AFaArticulo = " & CodArticulo _
                & " And AFaCategoriaD = CDtCatArticulo " _
                & " And CDtCatCliente = " & CodCatCliente _
                & " And CDtCatPlazo = " & plan
            
        Set RsBDC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        
        If Not RsBDC.EOF Then
            If Not IsNull(RsBDC!AFaCantidadD) Then
                If RsBDC!AFaCantidadD <= Cantidad Then
                    aRetorno = Unitario - (Unitario * RsBDC(0)) / 100
                    aRetorno = Redondeo(aRetorno, mMRound)
                Else
                    If MsgBox("La cantidad no llega a la mímima (" & RsBDC!AFaCantidadD & ") para aplicar el descuento. " & Chr(vbKeyReturn) _
                                & "Desea aplicar el descuento correspondiente.", vbQuestion + vbYesNo, Trim(Articulo)) = vbYes Then
                        aRetorno = Unitario - (Unitario * RsBDC(0)) / 100
                        aRetorno = Redondeo(aRetorno, mMRound)
                    End If
                End If
            End If
        End If
        
        RsBDC.Close
    End If
    BuscoDescuentoCliente = CStr(aRetorno)
    Exit Function

errDescuento:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al procesar los descuentos por cliente.", Err.Description
    BuscoDescuentoCliente = aRetorno
End Function

Private Sub InsertoFila()

    On Error GoTo ErrIF
    'Valido los campos para insertar la linea de articulo-----------------------------------------------------
    If cCuota.ListIndex = -1 Then
        MsgBox "Debe seleccionar el tipo de financiación.", vbExclamation, "ATENCIÓN"
        Foco cCuota: Exit Sub
    End If
    If cArticulo.ListIndex = -1 Then
        MsgBox "Debe seleccionar un artículo.", vbExclamation, "ATENCIÓN"
        Foco cArticulo: Exit Sub
    End If
    If Not IsNumeric(tCantidad.Text) Then
        MsgBox "La cantidad ingresada no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tCantidad: Exit Sub
    End If
    If Not Val(tCantidad.Text) > 0 Then
        MsgBox "La cantidad ingresada no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tCantidad: Exit Sub
    End If
    If Not IsNumeric(tUnitario.Text) Then
        MsgBox "El precio unitario ingresado no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tUnitario: Exit Sub
    Else
        If CCur(tUnitario.Text) < 0 Then
            MsgBox "No se puede facturar artículos con costo negativo.", vbExclamation, "ATENCIÓN"
            Foco tUnitario: Exit Sub
        End If
    End If
    
    If (Val(tUnitario.Text) - (Val(tUnitario.Text) \ 1)) <> 0 And cMoneda.ItemData(cMoneda.ListIndex) = paMonedaPesos Then
        If MsgBox("El valor de la cuota debe ser entero, no con decimales." & vbCrLf & "Desea continuar.", vbExclamation + vbYesNo + vbDefaultButton2, "Valor con Decimales") = vbNo Then
            Foco tUnitario: Exit Sub
        End If
    End If
    '-----------------------------------------------------------------------------------------------------------------------
    
    'Veo si es combo
    Dim bEsCombo As Boolean
    Cons = "Select * from Presupuesto, PresupuestoArticulo  " & _
               " Where PreArtCombo = " & cArticulo.ItemData(cArticulo.ListIndex) & _
               " And PreID = PArPresupuesto" & _
               " And PreMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then bEsCombo = True Else bEsCombo = False
    RsAux.Close
    
    If bEsCombo Then
        Dim aImporteD As Currency
        If Not sConEntrega Then aImporteD = CCur(lSubTotalF.Caption) Else aImporteD = CCur(tUnitario.Text)
        ProcesoArticuloCombo cArticulo.ItemData(cArticulo.ListIndex), bInsertarFilas:=True, ImporteDigitado:=aImporteD
    Else
    
        AgregoFila sConEntrega, cCuota.ItemData(cCuota.ListIndex), Val(tCantidad.Tag), cArticulo.ItemData(cArticulo.ListIndex), Val(tCantidad.Text), _
                        oArtEdicion.NombreArticulo, cCuota.Text, Trim(tUnitario.Tag), CCur(tUnitario.Text), oArtEdicion
    End If
    
    LimpioRenglon
    cArticulo.Clear: cArticulo.Tag = ""
    cCuota.SetFocus
    If lvVenta.ListItems.Count > 0 Then
        cMoneda.Enabled = False: MnuEmitir.Enabled = True
    Else
        cMoneda.Enabled = True: MnuEmitir.Enabled = False
    End If
    HabilitoEntrega
    Exit Sub
    
ErrIF:
    clsGeneral.OcurrioError "Ocurrió un error al insertar el renglón.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AgregoFila(bConEntrega As Boolean, lTCuota As Long, lPlan As Long, IDArticulo As Long, iQ As Long, tTxtArticulo As String, tTxtCuota As String, _
                                  cUnitarioF As String, cValorDigitado As Currency, ByVal artEditado As clsArticulo)

    Dim pstrKEY As String
    pstrKEY = IIf(bConEntrega, "E", "N") & lTCuota & "P" & lPlan & "A" & IDArticulo
    If Val(lArticulo.Tag) = TFactura.ArtEspecifico Then
        pstrKEY = pstrKEY & "F" & Val(cArticulo.Tag)
    Else
        pstrKEY = pstrKEY & "F0"
    End If
    
    Set itmX = lvVenta.ListItems.Add(, pstrKEY, Trim(tTxtCuota))
    
    If bConEntrega Then
        'Set itmX = lvVenta.ListItems.Add(, "E" & lTCuota & "P" & lPlan & "A" & lArticulo, Trim(tTxtCuota))
        itmX.SubItems(8) = cUnitarioF                           'Unitario Contado (para controlar cambio de precios con entrega)
        itmX.SubItems(3) = Format(cValorDigitado, "#,##0.00")   'Contado
    Else
        'Set itmX = lvVenta.ListItems.Add(, "N" & lTCuota & "P" & lPlan & "A" & lArticulo, Trim(tTxtCuota))
        itmX.SubItems(3) = Format(cUnitarioF, "#,##0.00")                     'Contado
    End If
    itmX.SubItems(1) = CStr(iQ)
    itmX.SubItems(2) = Trim(tTxtArticulo)
    itmX.SubItems(4) = IVAArticulo(IDArticulo)
        
    If Trim(lSubTotalF.Caption) <> "" Then
        'Cuota
        itmX.SubItems(6) = Format(cValorDigitado * iQ, FormatoMonedaP)
        
        'Ajusto el subtotal con lo que me da la cuota (SubTotal)
        itmX.SubItems(7) = lSubTotalF.Caption   'Total Financiado
        itmX.SubItems(8) = lSubTotalF.Tag     'Unitario Financiado
        TotalesSumo CCur(itmX.SubItems(7)), CCur(itmX.SubItems(4))
    End If
    colArtsGrilla.Add artEditado
    If (artEditado.VentaXMayor = 0 And InStr(1, prmCategoriaDistribuidor, "," & Val(labDireccion.Tag) & ",") > 0) Or _
        (artEditado.VentaXMayor > 1 And artEditado.VentaXMayor < CInt(itmX.SubItems(1))) Then
        itmX.ForeColor = &HFF&
    Else
        itmX.ForeColor = vbBlack
    End If

End Sub

Private Function IVAArticulo(lngCodigo As Long)

    IVAArticulo = 0
    Cons = "Select IVAPorcentaje From ArticuloFacturacion, TipoIva " _
        & " Where AFaArticulo = " & lngCodigo _
        & " And AFaIVA = IVACodigo"
        
    Set RsArt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not RsArt.EOF Then IVAArticulo = Format(RsArt(0), "#0.00")
    RsArt.Close

End Function


Private Function Ingresado(Articulo As Long, Cuota As Long)

    If lvVenta.ListItems.Count > 0 Then
        Ingresado = False
        For Each itmX In lvVenta.ListItems
            If ValorClave(itmX.Key, "C") = Cuota And ValorClave(itmX.Key, "A") = Articulo Then
                Ingresado = True
                Exit Function
            End If
        Next
    Else
        Ingresado = False
    End If

End Function

Private Sub TotalesResto(Total As Currency, Iva As Currency)

    labIVA.Caption = Format(CCur(labIVA.Caption) - (Total - (Total / (1 + Iva / 100))), "#,##0.00")
    labTotal.Caption = Format(CCur(labTotal.Caption) - Total, "#,##0.00")
    labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIVA.Caption), "#,##0.00")
    
End Sub

Private Sub TotalesSumo(Total As Currency, Iva As Currency)

    labIVA.Caption = Format(CCur(labIVA.Caption) + Total - (Total / (1 + Iva / 100)), "#,##0.00")
    labTotal.Caption = Format(CCur(labTotal.Caption) + Total, "#,##0.00")
    labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIVA.Caption), "#,##0.00")
        
End Sub

Private Sub tUsuario_GotFocus()

    tUsuario.SelStart = 0
    tUsuario.SelLength = Len(tUsuario.Text)
    Status.Panels(1).Text = " Ingrese el dígito de usuario."

End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And IsNumeric(tUsuario.Text) Then
        tUsuario.Tag = BuscoUsuario(Val(tUsuario.Text))
        If tUsuario.Tag = 0 Then
            tUsuario.Text = vbNullString
            tUsuario.Tag = vbNullString
            Exit Sub
        End If
        If tUsuario.Tag <> vbNullString Then cComentario.SetFocus
    End If
    
End Sub

Private Sub AccionGrabar()

    On Error GoTo errGrabar
    If Not ControloDatos Then Exit Sub
        
    If labDireccion.Tag = "" Then labDireccion.Tag = "0"
    
    Dim bSucPrecios As Boolean, bSucPlanNoHab As Boolean, bArtVtaXMayor As Boolean
    
    ControloPrecios CLng(labDireccion.Tag), bSucPrecios, bSucPlanNoHab, bArtVtaXMayor

    If bSucPlanNoHab Or bSucPrecios Or bArtVtaXMayor Then
        
        Dim mTipoSuceso As Integer
        If bSucPlanNoHab And bSucPrecios And bArtVtaXMayor Then
            aTexto = "Cambio de Precios/Planes No Habilitados/X Mayor Inhabilitado"
            mTipoSuceso = TipoSuceso.ModificacionDePrecios
        Else
            If (bSucPlanNoHab Or bArtVtaXMayor) Then
                If bSucPlanNoHab And bArtVtaXMayor Then
                    aTexto = "Solicitud con Plan No Habilitado/X Mayor inhabilitado"
                ElseIf bSucPlanNoHab Then
                    aTexto = "Solicitud con Plan No Habilitado"
                Else
                    aTexto = "Venta x mayor inhabilitada"
                End If
                mTipoSuceso = TipoSuceso.FacturaArticuloInhabilitado
            Else
                If Not bSucPlanNoHab And bSucPrecios Then
                    aTexto = "Cambio de Precios"
                    mTipoSuceso = TipoSuceso.ModificacionDePrecios
                End If
            End If
        End If
        
        gSucesoUsr = 0
        gSucesoUsrAut = 0
        Dim objSuceso As New clsSuceso
        objSuceso.TipoSuceso = mTipoSuceso
        objSuceso.ActivoFormulario CLng(tUsuario.Tag), aTexto, cBase
        
        Me.Refresh
        gSucesoUsr = objSuceso.RetornoValor(Usuario:=True)
        gSucesoDef = objSuceso.RetornoValor(Defensa:=True)
        gSucesoUsrAut = objSuceso.Autoriza
        Set objSuceso = Nothing
        
        If gSucesoUsr = 0 Or Trim(gSucesoDef) = "" Then Exit Sub  'Abortó el ingreso del suceso
        
    End If
    
    snd_ActivarSonido prmPathSound & sndGrabar
    If MsgBox("Confirma almacenar los datos ingresados en la solicitud.", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    
    GrabarSolicitud
    Exit Sub

errGrabar:
    clsGeneral.OcurrioError "Error al procesar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub ControloPrecios(CategoriaCliente As Long, bDifPrecios As Boolean, bPlanNoHab As Boolean, bArtVtaXMayor As Boolean)

Dim aItmx As ListItem
Dim RsCon As rdoResultset
Dim aUnitario As Currency
Dim aDiferencia As Currency     'Diferencia de Precios

Dim aKArticulo As Long, aKPlan As Long, aKTCuota As Long, bOK As Boolean
Dim aIdx As Integer
Dim sPlanesNO As String

    On Error GoTo errControl
    bDifPrecios = False: bPlanNoHab = False
    ReDim arrSucDP(0)
    Screen.MousePointer = 11
        
    'Armo Str con Tipos de cuotas Deshabilitados ----------------------------------------
    sPlanesNO = ","
    Cons = "Select * from TipoCuota Where TCuDeshabilitado = 'S'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        sPlanesNO = sPlanesNO & RsAux!TCuCodigo & ","
        RsAux.MoveNext
    Loop
    RsAux.Close
    '---------------------------------------------------------------------------------------------
    Dim bEnUso As Boolean
    
    For Each aItmx In lvVenta.ListItems
        bEnUso = True
        aKArticulo = ValorClave(aItmx.Key, "A")
        
        'Controlo si el articulo esta EN USO        --------------------------------------------------------------
        Cons = "Select * From Articulo Where ArtID = " & aKArticulo
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux.EOF Then
            If RsAux!ArtEnUso = 0 Then
                If Not ValidoComboOK(aKArticulo, ValorClave(aItmx.Key, "C")) Then
                    bEnUso = False
                    aIdx = UBound(arrSucDP) + 1
                    ReDim Preserve arrSucDP(aIdx)
                    With arrSucDP(aIdx)
                        .IDArticulo = aKArticulo
                        .TSuceso = TipoSuceso.FacturaArticuloInhabilitado
                        .DifPrecio = 0
                        .TextoPlan = Trim(aItmx.Text)
                    End With
                    bPlanNoHab = True
                End If
            End If
            
            If Not IsNull(RsAux("ArtEnVentaXMayor")) Then
                
                'If (RsAux("ArtEnVentaXMayor") = 0 Or (RsAux("ArtEnVentaXMayor") < CCur(aItmx.SubItems(1)) And RsAux("ArtEnVentaXMayor") > 1)) And InStr(1, prmCategoriaDistribuidor, "," & Val(labDireccion.Tag) & ",") > 0 Then
                Dim bAdd As Boolean
                
                bAdd = False
                If (RsAux("ArtEnVentaXMayor") = 0 And InStr(1, prmCategoriaDistribuidor, "," & Val(labDireccion.Tag) & ",") > 0) Then
                    bAdd = True
                    MsgBox "Atención!!! " & vbCrLf & vbCrLf & "No está autorizada la venta del artículo " & aItmx.SubItems(2) & " a distribuidores." & vbCrLf & vbCrLf & "Debe consultar para vender.", vbExclamation, "POSIBLE ERROR"
                ElseIf RsAux("ArtEnVentaXMayor") > 1 And RsAux("ArtEnVentaXMayor") < CInt(aItmx.SubItems(1)) Then
                    bAdd = True
                    MsgBox "Atención!!! " & vbCrLf & vbCrLf & "La cantidad máxima autorizada de venta para el artículo " & aItmx.SubItems(2) & " es de  " & RsAux("ArtEnVentaXMayor") & vbCrLf & vbCrLf & "Debe consultar para exceder dicha cantidad.", vbExclamation, "POSIBLE ERROR"
                End If
                
                If bAdd Then
                    aIdx = UBound(arrSucDP) + 1
                    ReDim Preserve arrSucDP(aIdx)
                    With arrSucDP(aIdx)
                        .IDArticulo = aKArticulo
                        .TSuceso = TipoSuceso.FacturaArticuloInhabilitado
                        .DifPrecio = 0
                        .TextoPlan = Trim(aItmx.Text)
                    End With
                    bArtVtaXMayor = True
                End If
                
            End If
            
        End If
        RsAux.Close
        
        If bEnUso Then      '---------------------------------------------------------------------------------------------------------------
            If Trim(Val(aItmx.SubItems(8))) <> 0 Then       'Si tenia precio para controlar
                aDiferencia = 0
                If Trim(aItmx.SubItems(5)) = "" Then    '(5)= Entrega -- Si el Plan es con entrega no hago el control
                    'SOLO PARA PLANES SIN ENTREGA!!!!!!
                    aUnitario = CCur(aItmx.SubItems(8))
                    If aUnitario * CCur(aItmx.SubItems(1)) <> CCur(aItmx.SubItems(7)) Then
                        '2 posibilidades --> o hay descuento o cambió el precio
                        'Veo Si hay descuentos
                        If CategoriaCliente > 0 Then        'Hago las consultas para comparar precios
                            aKArticulo = ValorClave(aItmx.Key, "A")
                            aKTCuota = ValorClave(aItmx.Key, "C")
                            
                            Cons = "Select CDTPorcentaje, AFaCantidadD, TCuCantidad From ArticuloFacturacion, CategoriaDescuento, TipoCuota" _
                                    & " Where AFaArticulo = " & aKArticulo _
                                    & " And AFaCategoriaD = CDtCatArticulo " _
                                    & " And CDtCatCliente = " & CategoriaCliente _
                                    & " And CDtCatPlazo = " & aKTCuota _
                                    & " And CDtCatPlazo = TCuCodigo "
                                
                            Set RsCon = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                            If Not RsCon.EOF Then
                                'Unitario - (Unitario * %Dto) / 100
                                aUnitario = Redondeo(CCur(aItmx.SubItems(8)) - (CCur(aItmx.SubItems(8)) * RsCon(0)) / 100, mMRound)
                                'Para que no de diferencia, hay que Unitario = Redondeo(Unitario/CCuotas)
                                aUnitario = CCur(Redondeo(aUnitario / RsCon!TCuCantidad, mMRound)) * RsCon!TCuCantidad
                            End If
                            RsCon.Close
                        End If
                        'OJO en el Sbt(7) está el financiado sin descuentos
                        'Si el unitario * cantidad <> Subtotal ------> HayQueGrabarSuceso
                        If aUnitario * CCur(aItmx.SubItems(1)) <> CCur(aItmx.SubItems(7)) Then
                            aDiferencia = CCur(aItmx.SubItems(7)) - aUnitario * CCur(aItmx.SubItems(1))
                        End If
                    End If
                Else        'Plan Con entrega Comparo si modificó el contado
                    aUnitario = CCur(aItmx.SubItems(8))     'Initario
                    If aUnitario <> CCur(aItmx.SubItems(3)) Then
                        aDiferencia = CCur(aItmx.SubItems(3)) - aUnitario
                    End If
                End If
                
                If aDiferencia <> 0 Then
                    aIdx = UBound(arrSucDP) + 1
                    ReDim Preserve arrSucDP(aIdx)
                    With arrSucDP(aIdx)
                        .IDArticulo = aKArticulo
                        .TSuceso = TipoSuceso.ModificacionDePrecios
                        .DifPrecio = aDiferencia
                        .TextoPlan = Trim(aItmx.Text)
                    End With
                    bDifPrecios = True
                End If
            End If
            
            If Trim(aItmx.SubItems(5)) = "" Then    '(5)= Entrega -- Si el Plan es con entrega no hago el control
                'Valido planes habilitados
                'Si no hay precios financiados y el coeficiente <> 1 va suceso por Plan No Habilitado
                'Si no hay ni Contado, ni Credito --> No va Suceso
                aKArticulo = ValorClave(aItmx.Key, "A")
                aKTCuota = ValorClave(aItmx.Key, "C")
                bOK = False
            
                If InStr(sPlanesNO, "," & aKTCuota & ",") <> 0 Then
                    aIdx = UBound(arrSucDP) + 1
                    ReDim Preserve arrSucDP(aIdx)
                    With arrSucDP(aIdx)
                        .IDArticulo = aKArticulo
                        .TSuceso = TipoSuceso.FacturaPlanInhabilitado
                        .DifPrecio = 0
                        .TextoPlan = Trim(aItmx.Text)
                    End With
                    bPlanNoHab = True
                Else
                    'Saco el valor de la cuota financiado
                    Cons = "Select PViPrecio, PViHabilitado, PViPlan From PrecioVigente" _
                            & " Where PVIArticulo = " & aKArticulo _
                            & " And PViMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                            & " And PViTipoCuota = " & aKTCuota _
                            & " And PViHabilitado = 1"
                    Set RsCon = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    If Not RsCon.EOF Then bOK = True
                    RsCon.Close
                    
                    If Not bOK Then
                        aKPlan = 0
                        'Saco el plan por el precio contado
                        Cons = "Select PViPrecio, PViHabilitado, PViPlan From PrecioVigente" _
                                & " Where PVIArticulo = " & aKArticulo _
                                & " And PViMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                                & " And PViTipoCuota = " & paTipoCuotaContado _
                                & " And PViHabilitado = 1"
                        Set RsCon = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                        If Not RsCon.EOF Then aKPlan = RsCon!PViPlan
                        RsCon.Close
                        
                        If aKPlan > 0 Then
                            bOK = True
                            'Busco el coeficiente p/tipo de Cuota y plan
                            Cons = "Select * from Coeficiente" & _
                                        " Where CoePlan = " & aKPlan & _
                                        " And CoeTipoCuota = " & aKTCuota & _
                                        " And CoeMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
                            
                            Set RsCon = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                            If Not RsCon.EOF Then
                                If RsCon!CoeCoeficiente = 1 Then bOK = False     'NO Va suceso
                            End If
                            RsCon.Close
                        End If
                        
                        If bOK Then     'Va suceso
                            aIdx = UBound(arrSucDP) + 1
                            ReDim Preserve arrSucDP(aIdx)
                            With arrSucDP(aIdx)
                                .IDArticulo = aKArticulo
                                .TSuceso = TipoSuceso.FacturaPlanInhabilitado
                                .DifPrecio = 0
                                .TextoPlan = Trim(aItmx.Text)
                            End With
                            bPlanNoHab = True
                        End If
                    End If
                End If
            End If
        End If
    Next
    Screen.MousePointer = 0
    Exit Sub

errControl:
    clsGeneral.OcurrioError "Error al controlar los precios de artículos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function ControloDatos() As Boolean
Dim Suma As Currency

    ControloDatos = False
    
    If txtCliente.Cliente.Codigo <= 0 Then
        MsgBox "No se puede emitir una solicitud sin seleccionar un cliente.", vbExclamation, "ATENCIÓN"
        Foco txtCliente: Exit Function
    End If
    If cMoneda.ListIndex = -1 Then
        MsgBox "Se debe seleccionar una moneda para emitir la solicitud.", vbExclamation, "ATENCIÓN"
        cMoneda.Enabled = True: Foco cMoneda: Exit Function
    End If
    If lvVenta.ListItems.Count = 0 Then
        MsgBox "Debe ingresar los artículos para la solicitud.", vbExclamation, "ATENCIÓN"
        Foco cArticulo: Exit Function
    End If
    
    If txtGarantia.Cliente.Codigo > 0 And txtGarantia.Cliente.Tipo = TC_Persona And txtGarantia.Cliente.Documento = "" Then
        MsgBox "La garantía seleccionada debe presentar la cédula de identidad.", vbExclamation, "ATENCIÓN"
        Foco txtGarantia: Exit Function
    End If
    
    'Valido la suma de los Subtotales contra el Total -----------------------------
    Suma = 0
    For Each itmX In lvVenta.ListItems
        If Trim(itmX.SubItems(7)) = "" Then
            MsgBox "Se debe ingresar el valor de la entrega para las financiaciones con entrega.", vbExclamation, "ATENCIÓN"
            Foco tEntregaT: Exit Function
        End If
        
        If CCur(itmX.SubItems(7)) = 0 And Val(lArticulo.Tag) <> TFactura.Servicio Then
            MsgBox "No se pueden solicitar artículos con importe cero.", vbExclamation, "ATENCIÓN"
            If cArticulo.Enabled Then cArticulo.SetFocus
            Exit Function
        End If
        
        Suma = Suma + CCur(itmX.SubItems(7))
    Next
    
    If Suma <> CCur(labTotal.Caption) Then
        MsgBox "La suma total no coincide con la suma de la lista, verifique.", vbCritical, "ATENCIÓN"
        cArticulo.SetFocus: Exit Function
    End If
    '-----------------------------------------------------------------------------------
    'Valido la suma de las Entregas con el Monto de Entrega
    Suma = -1
    For Each itmX In lvVenta.ListItems
        If Trim(itmX.SubItems(5)) <> "" Then    'Plan con entrega
            If Suma = -1 Then Suma = 0
            Suma = Suma + CCur(itmX.SubItems(5))
        End If
    Next
    If Suma <> -1 Then          'Hay entregas
        If Trim(tEntregaT.Text) <> "" Then
            If Suma <> CCur(tEntregaT.Text) Then
                MsgBox "La suma de las entregas no coincide con el valor ingresado, verifique.", vbExclamation, "ATENCIÓN"
                Foco tEntregaT: Exit Function
            End If
        Else
            MsgBox "La suma de las entregas no coincide con el valor ingresado, verifique.", vbExclamation, "ATENCIÓN"
            Foco tEntregaT: Exit Function
        End If
    End If
    '-----------------------------------------------------------------------------------
    
    If cPago.ListIndex = -1 Then
        MsgBox "Se debe seleccionar la forma de pago de la solicitud.", vbExclamation, "ATENCIÓN"
        Foco cPago: Exit Function
    End If
    
    If Val(tVendedor.Tag) = 0 Then 'Or tVendedor.Tag = "0" Or tVendedor.Tag = 0 Then
        MsgBox "Debe ingresar el dígito del vendedor.", vbExclamation, "ATENCIÓN"
        Foco tVendedor: Exit Function
    End If
    
    If tUsuario.Tag = vbNullString Then
        MsgBox "Debe ingresar el dígito de usuario.", vbExclamation, "ATENCIÓN"
        Foco tUsuario: Exit Function
    End If
    
    'Si el pago es con Cheque Dif. controlo que las cuotas no superen el parametro = paCantidadMaxCheques
    If cPago.ItemData(cPago.ListIndex) = TipoPagoSolicitud.ChequeDiferido Then
        For Each itmX In lvVenta.ListItems
            If CCur(itmX.SubItems(6)) > 0 Then
                If Trim(itmX.SubItems(5)) = "" Then
                    'Subtotal / Valor Cuota = Cant. Cuotas
                    If CCur(itmX.SubItems(7)) / CCur(itmX.SubItems(6)) > paCantidadMaxCheques Then
                        MsgBox "Hay cuotas que superan la cantidad máxima de cheques diferidos." & Chr(vbKeyReturn) & "Cambie la forma de pago.", vbExclamation, "ATENCIÓN"
                        Foco cPago: Exit Function
                    End If
                Else
                    '(Subtotal - Entrega) / Valor Cuota = Cant. Cuotas
                    If (CCur(itmX.SubItems(7)) - CCur(itmX.SubItems(5))) / CCur(itmX.SubItems(6)) > paCantidadMaxCheques Then
                        MsgBox "Hay cuotas que superan la cantidad máxima de cheques diferidos." & Chr(vbKeyReturn) & "Cambie la forma de pago.", vbExclamation, "ATENCIÓN"
                        Foco cPago: Exit Function
                    End If
                End If
            End If
        Next
    End If
    
    Dim aQTel As Integer
    If Not ValidoTelefonos(txtCliente.Cliente.Codigo, aQTel) Then Exit Function
    'Controlo cantidades de teléfonos ingresados ------------------------------------------------------
    If txtCliente.Cliente.Tipo = TC_Persona Then
        If aQTel < paQTelefonos Then
            If MsgBox("El cliente tiene menos de " & paQTelefonos & " teléfonos ingresados." & vbCrLf & _
                        "Quiere ingresar el teléfono del Trabajo o el de un Familiar ...", vbQuestion + vbYesNo, "Ingresar otros teléfonos") = vbYes Then
                Dim objCl As New clsCliente
                objCl.Personas idCliente:=txtCliente.Cliente.Codigo
                Me.Refresh
                Set objCl = Nothing
            End If
        End If
    End If
    '------------------------------------------------------------------------------------------------------------
    
    'Valdio Q de Artículos por Plan    ------------------------------------------------------------------------
    If prmQMaxArticulosPlan <> 0 And lvVenta.ListItems.Count > prmQMaxArticulosPlan Then
        Dim mYaOK As String, iJR As Integer, iQ As Integer, mTipoCta As Integer
        For iJR = 1 To lvVenta.ListItems.Count
            iQ = 0
            mTipoCta = ValorClave(lvVenta.ListItems(iJR).Key, "C")
            If InStr(mYaOK, "|" & mTipoCta & "|") = 0 Then
                For Each itmX In lvVenta.ListItems
                    If mTipoCta = ValorClave(itmX.Key, "C") Then iQ = iQ + 1
                Next
                mYaOK = mYaOK & "|" & mTipoCta & "|"
                
                If iQ > prmQMaxArticulosPlan Then
                    If MsgBox("Ud. ingresó más renglones de los que entran en el papel de la factura." & vbCrLf & "Se pueden ingresar hasta " & prmQMaxArticulosPlan & " por plan." & vbCrLf & vbCrLf & "Si continúa, los renglones no saldrán impresos." & vbCrLf & "¿Quiere continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "Demasiados Renglones") = vbNo Then
                        ControloDatos = False: Exit Function
                    End If
                End If
            
            End If
        Next
    End If
    '------------------------------------------------------------------------------------------------------------------------------------------------
    ControloDatos = True
    
    'Si paso la Validación controlo la direccion que factura-----------------------------------------------------------------
    If cDireccion.ListIndex <> -1 Then
        On Error Resume Next
        If gDirFactura <> cDireccion.ItemData(cDireccion.ListIndex) Then        'Cambio Dir Facutua
            If MsgBox("Ud. a cambiado la dirección con la que el cliente factura habitualmente." & vbCrLf & "Quiere que esta dirección quede por defecto para facturar.", vbQuestion + vbYesNo, "Dirección por Defecto al Facturar") = vbNo Then Exit Function
            
            If cDireccion.ItemData(cDireccion.ListIndex) <> Val(cDireccion.Tag) Then        'Dir. selecc. <> a la Ppal.
                
                Dim rsAD As rdoResultset
                Cons = "Select * from DireccionAuxiliar Where DAuCliente = " & txtCliente.Cliente.Codigo & " And DAuDireccion = " & cDireccion.ItemData(cDireccion.ListIndex)
                Set rsAD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not rsAD.EOF Then
                    rsAD.Edit: rsAD!DAuFactura = True: rsAD.Update
                End If
                rsAD.Close
            End If
            
            If gDirFactura <> Val(cDireccion.Tag) Then      'La gDirFactura Anterior no era la ppal, la desmarco
                Cons = "Select * from DireccionAuxiliar Where DAuCliente = " & txtCliente.Cliente.Codigo & " And DAuDireccion = " & gDirFactura
                Set rsAD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not rsAD.EOF Then
                    rsAD.Edit: rsAD!DAuFactura = False: rsAD.Update
                End If
                rsAD.Close
            End If
            gDirFactura = cDireccion.ItemData(cDireccion.ListIndex)
        End If
        
    End If
    '----------------------------------------------------------------------------------------------------------------------------------
        
End Function

Private Function BuscoUsuario(intUsuario As Integer) As Integer
On Error GoTo ErrBU

    Cons = "SELECT * FROM USUARIO WHERE UsuDigito = " & intUsuario
    Set RsArt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If RsArt.EOF Then
        BuscoUsuario = 0
        MsgBox "No existe un usuario con ese digito.", vbExclamation, "ATENCIÓN"
    Else
        BuscoUsuario = RsArt!UsuCodigo
    End If
    RsArt.Close
    Exit Function
    
ErrBU:
    clsGeneral.OcurrioError "Ocurrio un error inesperado."
    BuscoUsuario = 0
    
End Function

Private Sub GrabarSolicitud()

Dim aSolicitud As Long
Dim aIdCliente As Long

    Screen.MousePointer = vbHourglass
    On Error GoTo ErrGFR
    FechaDelServidor
    cBase.BeginTrans
    
    On Error GoTo ErrResumo
    aIdCliente = txtCliente.Cliente.Codigo
    
    'Cargo los datos de la SOLICITUD-------------------------------------------------------------------------------------
    Cons = "Select * From Solicitud Where SolCodigo = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.AddNew
    
    RsAux!SolCliente = aIdCliente
    RsAux!SolFecha = Format(gFechaServidor, sqlFormatoFH)
    
    Select Case Val(lArticulo.Tag)
        Case TFactura.Articulo, TFactura.ArtEspecifico: RsAux!SolTipo = TipoSolicitud.AlMostrador
        Case TFactura.Servicio: RsAux!SolTipo = TipoSolicitud.Servicio: RsAux!SolIdServicio = Val(cArticulo.Text)
    End Select
        
    RsAux!SolProceso = TipoResolucionSolicitud.Manual
    RsAux!SolEstado = EstadoSolicitud.Pendiente
    
    If txtGarantia.Cliente.Codigo > 0 Then RsAux!SolGarantia = txtGarantia.Cliente.Codigo
    RsAux!SolFormaPago = cPago.ItemData(cPago.ListIndex)
    RsAux!SolMoneda = cMoneda.ItemData(cMoneda.ListIndex)
    If Trim(cComentario.Text) <> "" Then RsAux!SolComentarioS = Trim(cComentario.Text)
    RsAux!SolUsuarioS = tUsuario.Tag
    RsAux!SolSucursal = paCodigoDeSucursal
    RsAux!SolVendedor = CLng(tVendedor.Tag)
    
    If prmIDLlamada > 0 Then RsAux!SolLlamada = prmIDLlamada
    
    RsAux.Update
    RsAux.Close
    '---------------------------------------------------------------------------------------------------------------------------
    
    'Saco el numero de la solicitud--------------------------------------------------
    Cons = "SELECT MAX(SolCodigo) From Solicitud Where SolCliente = " & aIdCliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    aSolicitud = RsAux(0)
    RsAux.Close
    '-------------------------------------------------------------------------------------
    
    'Inserto los Renglones de la Solicitud--------------------------------------------------------
    Cons = "Select * From RenglonSolicitud Where RSoSolicitud = " & aSolicitud
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    For Each itmX In lvVenta.ListItems
        RsAux.AddNew
        
        RsAux!RSoSolicitud = aSolicitud
        RsAux!RSoTipoCuota = ValorClave(itmX.Key, "C")
        RsAux!RSoArticulo = ValorClave(itmX.Key, "A")
        If Trim(itmX.SubItems(5)) <> "" Then RsAux!RSoValorEntrega = CCur(itmX.SubItems(5))
        RsAux!RSoValorCuota = CCur(itmX.SubItems(6))
        RsAux!RSoCantidad = itmX.SubItems(1)
        
        RsAux.Update
        
        'Up
        If InStr(1, itmX.Key, "F", vbTextCompare) > 0 Then
            If ValorClave(itmX.Key, "F") > 0 Then
                Cons = "Update ArticuloEspecifico Set AEsTipoDocumento = 2," & _
                                                    " AEsDocumento = " & aSolicitud & "," & _
                                                    " AEsModificado = '" & Format(gFechaServidor, "yyyy/mm/dd hh:nn:ss ") & "'" & _
                        " Where AEsID = " & ValorClave(itmX.Key, "F")
                cBase.Execute Cons
            End If
        End If
    Next
    RsAux.Close
    '-------------------------------------------------------------------------------------------------
        
    cBase.CommitTrans                               'Fin TRANSACCION----------------------------------------------!!!!!!!!!!!!!!!!!!!
    
    On Error GoTo errSucesos
    For i = 1 To UBound(arrSucDP)
        aTexto = "Solicitud de Crédito Nº " & aSolicitud & " (" & Trim(arrSucDP(i).TextoPlan) & ")"
        clsGeneral.RegistroSucesoAutorizado cBase, gFechaServidor, arrSucDP(i).TSuceso, paCodigoDeTerminal, gSucesoUsr, 0, Articulo:=arrSucDP(i).IDArticulo, _
                             Descripcion:=aTexto, Defensa:=Trim(gSucesoDef), _
                             Valor:=arrSucDP(i).DifPrecio, idCliente:=aIdCliente, idAutoriza:=gSucesoUsrAut
    Next
    
    txtCliente.Text = ""
    LimpioTodaLaFicha True
    prmIDLlamada = 0
    code_ProcesoScript "SOL10", aIdCliente, aSolicitud
    
    fnc_CallResolAutomatica aSolicitud
    
    Screen.MousePointer = 0
    Exit Sub
    
ErrGFR:
    clsGeneral.OcurrioError "Error al iniciar la transacción.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
ErrResumo:
    Resume ErrRelajo
ErrRelajo:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Error al realizar la solicitud.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
    
errSucesos:
    clsGeneral.OcurrioError "Error al grabar los sucesos.", Err.Description
    txtCliente.Text = ""
    LimpioTodaLaFicha True
    Screen.MousePointer = 0
End Sub

Private Function ValidoBuscarArticulo()

    ValidoBuscarArticulo = False
    
    If cMoneda.ListIndex = -1 Then
        MsgBox "Se debe seleccionar una moneda para realizar la solicitud.", vbCritical, "ATENCIÓN"
        Foco cMoneda: Exit Function
    End If

    If cCuota.ListIndex = -1 Then
        MsgBox "Se debe seleccionar el tipo de financiación para realizar la solicitud.", vbCritical, "ATENCIÓN"
        Foco cCuota: Exit Function
    End If
    
    ValidoBuscarArticulo = True

End Function

'----------------------------------------------------------------------------------------------------------------------------
Private Sub DistribuirEntregas(ValorAEntregar As Currency)

Dim sHay As Boolean
Dim iAuxiliar As Currency
Dim aTotal As Currency

    On Error GoTo errDistribuir
    sHay = False
    'Verifico si hay Cuotas con Entegas-------------------------------------------------------------------------
    For Each itmX In lvVenta.ListItems
        If Left(itmX.Key, 1) = "E" Then
            sHay = True: Exit For
        End If
    Next
    If Not sHay Then
        Screen.MousePointer = 0
        MsgBox "No hay financiaciones con entrega para realizar la distribución.", vbInformation, "ATENCIÓN"
        tEntregaT.Text = "": Exit Sub
    End If
    '-----------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    'Limpio los campos para recalcular Y saco el Total de precios con entrega-----------
    For Each itmX In lvVenta.ListItems
        If Left(itmX.Key, 1) = "E" Then
            
            If Trim(itmX.SubItems(7)) <> "" Then TotalesResto CCur(itmX.SubItems(7)), CCur(itmX.SubItems(4))
            
            itmX.SubItems(5) = ""
            itmX.SubItems(6) = ""
            itmX.SubItems(7) = ""
            aTotal = aTotal + (CCur(itmX.SubItems(3)) * CCur(itmX.SubItems(1)))
        End If
    Next
    
    If aTotal <= ValorAEntregar Then
        Screen.MousePointer = 0
        MsgBox "El valor a entregar no puede superar los valores contado de los artículos. Verifique los datos.", vbExclamation, "ATENCIÓN"
        Foco tEntregaT: Exit Sub
    End If
    '--------------------------------------------------------------------------------------------------
    
    Dim aDfs As Currency: aDfs = 0
    Dim itmP As ListItem
    For Each itmX In lvVenta.ListItems
        If Left(itmX.Key, 1) = "E" Then
            'Veo Si ya hice la distribucion
            If Trim(itmX.SubItems(5)) = "" Then
                'Con el Total distriubuyo el porcentaje de la entrega
                For Each itmP In lvVenta.ListItems
                    If itmX.Text = itmP.Text Then   'El mismo Tipo de Cuota
                                                
                        'Cambio del 9/8/01
                        'itmP.SubItems(5) = Format(((CCur(itmP.SubItems(3)) * CCur(itmP.SubItems(1)) * 100) / aTotal) * ValorAEntregar / 100, "#,##0")
                        itmP.SubItems(5) = Redondeo(((CCur(itmP.SubItems(3)) * CCur(itmP.SubItems(1)) * 100) / aTotal) * ValorAEntregar / 100, mMRound)
                        itmP.SubItems(5) = Format(itmP.SubItems(5), "#,##0.00")
                        
                        aDfs = aDfs + CCur(itmP.SubItems(5)) - Format(((CCur(itmP.SubItems(3)) * CCur(itmP.SubItems(1)) * 100) / aTotal) * ValorAEntregar / 100, "#,##0.00")
                        If InStr(CStr(aDfs), ".") = 0 And aDfs <> 0 Then
                            itmP.SubItems(5) = Format(CCur(itmP.SubItems(5)) - aDfs, "#,##0.00")
                            aDfs = 0
                        End If
                                                
                        'El valor de la cuota es el (Precio Contado - Entrega) * Coeficiente ----- Coeficiente (Plan, TCuota, Moneda)
                        Cons = "Select * from Coeficiente, TipoCuota" _
                            & " Where CoePlan = " & ValorClave(itmP.Key, "P") _
                            & " And CoeTipoCuota = " & ValorClave(itmP.Key, "C") _
                            & " And CoeMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                            & " And CoeTipocuota = TCuCodigo"
                        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                        'Calculo lo que queda por pagar ((P.contado * Cantidad)- Entega * Coeficiente)
                        iAuxiliar = ((CCur(itmP.SubItems(3)) * CCur(itmP.SubItems(1))) - CCur(itmP.SubItems(5))) * RsAux!CoeCoeficiente
                        
                        'Veo si tiene descuento
                        iAuxiliar = CCur(BuscoDescuentoCliente(ValorClave(itmP.Key, "A"), Val(labDireccion.Tag), iAuxiliar, CCur(itmP.SubItems(1)), itmP.SubItems(2), ValorClave(itmP.Key, "C")))
                        
                        'Valor de Cada Cuota
                        iAuxiliar = Redondeo(iAuxiliar / RsAux!TCuCantidad, mMRound)
                        itmP.SubItems(6) = Format(iAuxiliar, "#,##0.00")
                        'SubTotal = (Entrega + Las cuotas)
                        itmP.SubItems(7) = Format((CCur(itmP.SubItems(6)) * RsAux!TCuCantidad) + CCur(itmP.SubItems(5)), "#,##0.00")
                        
                        TotalesSumo CCur(itmP.SubItems(7)), CCur(itmP.SubItems(4))
                        RsAux.Close
                    End If
                Next
            End If
        End If
    Next
    Screen.MousePointer = 0
    Exit Sub

errDistribuir:
    clsGeneral.OcurrioError "Error al realizar la distribución de la entrega.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub HabilitoEntrega()

Dim sHay As Boolean

    sHay = False
    For Each itmX In lvVenta.ListItems
        If Left(itmX.Key, 1) = "E" Then
            sHay = True: Exit For
        End If
    Next
    
    If sHay Then
        tEntregaT.Enabled = True: tEntregaT.BackColor = Obligatorio
    Else
        tEntregaT.Enabled = False: tEntregaT.BackColor = Inactivo
    End If
    
End Sub

Private Sub IngresoDeEntrega(Valor As Boolean)

    TipoFacturacion TFactura.Articulo     'Por defecto siempre facturo articulos
    
    cCuota.Enabled = Not Valor
    cArticulo.Enabled = Not Valor
    tCantidad.Enabled = Not Valor
    tUnitario.Enabled = Not Valor
    
    tEntrega.Enabled = Valor
    
    If Valor Then
        cCuota.BackColor = Inactivo
        cArticulo.BackColor = Inactivo
        tCantidad.BackColor = Inactivo
        tUnitario.BackColor = Inactivo
        tEntrega.BackColor = Obligatorio
    Else
        cCuota.BackColor = Obligatorio
        cArticulo.BackColor = Obligatorio
        tCantidad.BackColor = Obligatorio
        tUnitario.BackColor = Obligatorio
        tEntrega.BackColor = Inactivo
    End If
        
End Sub

Private Sub ValidoMayorDeEdad(Titular As Boolean, Garantia As Boolean)

    If Titular Then
        If Trim(lTEdad.Caption) = "" Then
            Screen.MousePointer = 0
            If MsgBox("El cliente seleccionado no tiene ingresada la fecha de nacimiento." & vbCrLf & "Para realizar la solicitud debe ingresarla, desea hacerlo.", vbQuestion + vbYesNo, "Falta F/Nacimiento") = vbNo Then
                LimpioDatosCliente
            Else
                txtCliente.Comportamiento = CNC_Editar
                txtCliente.AbrirMantenimiento False
                txtCliente.Comportamiento = CNC_SiDatosCredito
            End If
        Else
            If CLng(lTEdad.Caption) < paMayorDeEdad Then
                Screen.MousePointer = 0
                MsgBox "El cliente seleccionado es menor de edad. No puede solicitar créditos.", vbExclamation, "ATENCIÓN"
                LimpioDatosCliente
            End If
        End If
    Else
        If Val(lGEdad.Tag) = 1 Then
            If Trim(lGEdad.Caption) = "" Then
                Screen.MousePointer = 0
                If MsgBox("El cliente seleccionado no tiene ingresada la fecha de nacimiento. Para actuar como garante debe ingresarla, desea hacerlo.", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
                    txtGarantia.Text = ""
                    lGarantia.Caption = ""
                    lGEdad.Caption = ""
                Else
                    txtGarantia.Comportamiento = CNC_Editar
                    txtGarantia.AbrirMantenimiento False
                    txtGarantia.Comportamiento = CNC_SiDatosCredito
                End If
            Else
                If CLng(lGEdad.Caption) < paMayorDeEdad Then
                    Screen.MousePointer = 0
                    MsgBox "El cliente seleccionado es menor de edad. No puede presentarse como garantía del crédito.", vbExclamation, "ATENCIÓN"
                    txtGarantia.Text = ""
                    lGarantia.Caption = ""
                    lGEdad.Caption = ""
                End If
            End If
        End If
    End If
    
End Sub

Private Function ValidoTelefonos(idCliente As Long, Optional aQTelef As Integer = 0) As Boolean

    On Error GoTo errTel
    ValidoTelefonos = True
    
    'Valido si tiene ingresado telefonos en la BD-------------------------------------------------------------
    Screen.MousePointer = 11
    Dim rsTel As rdoResultset, bHayTel As Boolean
    aQTelef = 0: bHayTel = False
    Cons = "Select Count(*) from Telefono Where TelCliente = " & idCliente
    Set rsTel = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsTel.EOF Then
        If Not IsNull(rsTel(0)) Then
            If rsTel(0) > 0 Then bHayTel = True
            aQTelef = rsTel(0)
        End If
    End If
    rsTel.Close
    
    Screen.MousePointer = 0
    
    If Not bHayTel Then
        MsgBox "Para solicitar un crédito es necesario ingresar los teléfonos del cliente." & vbCrLf & _
                    "Presione 'Aceptar' y luego [F2] para ingresar los datos que faltan.", vbInformation, "Faltan Teléfonos"
        ValidoTelefonos = False
    End If
    '--------------------------------------------------------------------------------------------------------------
errTel:
End Function

Private Sub tVendedor_Change()
    tVendedor.Tag = 0
End Sub

Private Sub tVendedor_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn And IsNumeric(tVendedor.Text) Then
        tVendedor.Tag = BuscoUsuario(Val(tVendedor.Text))
        If tVendedor.Tag = 0 Then
            tVendedor.Text = vbNullString
            tVendedor.Tag = vbNullString
            Exit Sub
        End If
        If tVendedor.Tag <> vbNullString Then Foco tUsuario
    End If
    
End Sub

Private Function ValorClave(Clave As String, Valor As String) As Long
'A=Artículo;    C=TipodeCuota;    P=Plan;   F=Art.Específico

    'E/N & IDTCTA & P & IDPLAN & A & IDART & F & IDAESP
    Select Case Valor
        Case "A":
            If InStr(1, Clave, "F", vbTextCompare) > 0 Then
                ValorClave = Mid(Clave, InStr(Clave, "A") + 1, InStr(Clave, "F") - InStr(Clave, "A") - 1)
            Else
                ValorClave = Mid(Clave, InStr(Clave, "A") + 1)
            End If
                 'ValorClave = CLng(Mid(Clave, InStr(Clave, "A") + 1, Len(Clave)))
        
        Case "C": ValorClave = CLng(Mid(Clave, 2, InStr(Clave, "P") - 2))
        Case "P": ValorClave = Mid(Clave, InStr(Clave, "P") + 1, InStr(Clave, "A") - InStr(Clave, "P") - 1)
        
        Case "F": ValorClave = CLng(Mid(Clave, InStr(Clave, "F") + 1, Len(Clave)))
    End Select
    
End Function

Private Sub CargoDatosServicio(IdServicio As Long)
    
    On Error GoTo ErrCDS
    Screen.MousePointer = 11
    Dim rsSer As rdoResultset
    
    lvVenta.ListItems.Clear         'OJO HAY Que Limpiar todo
    Set colArtsGrilla = New Collection
    labTotal.Caption = "0.00": labIVA.Caption = "0.00": labSubTotal.Caption = "0.00"
    
    Cons = "Select * From Servicio Where SerCodigo = " & IdServicio & " And SerDocumento Is Null"
    Set rsSer = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    If rsSer.EOF Then
        MsgBox "No se encontraron datos para ese servicio Nº " & Trim(cArticulo.Text) & ", o ya fue facturado.", vbExclamation, "ATENCIÓN"
        rsSer.Close
        Screen.MousePointer = 0: Exit Sub
    End If
    
    If Not IsNull(rsSer!SerMoneda) Then BuscoCodigoEnCombo cMoneda, rsSer!SerMoneda
    cMoneda.Enabled = False
    
    Dim aTotalContado As Currency: aTotalContado = 0
    Dim rsSR As rdoResultset
    Cons = "Select * From ServicioRenglon, Articulo " _
            & " Where SReServicio = " & IdServicio _
            & " And SReTipoRenglon = " & 2 & " And SReMotivo = ArtID"       'TipoRenglonS.Cumplido
    Set rsSR = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not rsSR.EOF
        InsertoArticuloServicio rsSR!ArtID, rsSR!ArtNombre, rsSR!SReTotal, rsSR!SReCantidad
        aTotalContado = aTotalContado + (Format(rsSR!SReTotal, FormatoMonedaP) * rsSR!SReCantidad)
        rsSR.MoveNext
    Loop
    rsSR.Close
    
    'Ahora en base al valor del costo final verifico el total que me dio los artículos.
    If aTotalContado <> rsSer!SerCostoFinal Then
        
        Cons = "Select * from Articulo Where ArtID = " & paArticuloCobroServicio
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
        If Not RsAux.EOF Then aTexto = RsAux!ArtNombre Else aTexto = ""
        RsAux.Close
        
        'Como tengo diferencia en la suma de los artículos con el costo final inserto artículo que me da el costo final.
        InsertoArticuloServicio paArticuloCobroServicio, aTexto, rsSer!SerCostoFinal - aTotalContado, 1
    End If
    
    HabilitoEntrega
    If lvVenta.ListItems.Count > 0 Then
        cCuota.Enabled = False: cCuota.BackColor = Inactivo
        cArticulo.Enabled = False: cArticulo.BackColor = Inactivo
        If tEntregaT.Enabled Then Foco tEntregaT Else Foco cPago
    End If
    
    rsSer.Close
    Screen.MousePointer = 0
    Exit Sub

ErrCDS:
    clsGeneral.OcurrioError "Error al buscar el servicio para facturar.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub InsertoArticuloServicio(aIdArticulo As Long, Articulo As String, PrecioUnitario As Currency, Cantidad As Currency)

Dim rsSR As rdoResultset
Dim aIDPlan As Long, aIDCoeficiente, aUnitario As Currency
Dim mIDMoneda As Long

    mIDMoneda = cMoneda.ItemData(cMoneda.ListIndex)
        
    'Busco el plan para el tipo de cuota seleccionado seleccionado-------------------------
    Cons = "Select PViPrecio, PViHabilitado, PViPlan From PrecioVigente" _
            & " Where PVIArticulo = " & aIdArticulo _
            & " And PViMoneda = " & mIDMoneda _
            & " And PViTipoCuota = " & paTipoCuotaContado _
            & " And PViHabilitado = 1"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then aIDPlan = RsAux!PViPlan Else aIDPlan = paPlanPorDefecto
    RsAux.Close
        
    'Valido que exista un coeficiente para el calculo (TipoCuota, Plan, Moneda) Ya sea para el plan ingresado o el por defecto
    Cons = "Select * from Coeficiente" _
            & " Where CoePlan = " & aIDPlan _
            & " And CoeTipoCuota = " & cCuota.ItemData(cCuota.ListIndex) _
            & " And CoeMoneda = " & mIDMoneda
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If RsAux.EOF Then    'Si NO hay coeficientes NO SE VENDE
        aIDCoeficiente = 1
        MsgBox "No existe un coeficiente para el cálculo de cuotas. Consulte.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0: cArticulo.Clear
    Else
        aIDCoeficiente = RsAux!CoeCoeficiente
    End If
    RsAux.Close
        
    'INSERTO LISTA A ARTICULOS------------------------------------------------------------------------------------------------------------------------
    If sConEntrega Then
        Set itmX = lvVenta.ListItems.Add(, "E" & cCuota.ItemData(cCuota.ListIndex) & "P" & aIDPlan & "A" & aIdArticulo, cCuota.Text)
        itmX.SubItems(8) = PrecioUnitario    'Unitario Contado (para controlar cambio de precios con entrega)
        itmX.SubItems(3) = Format(PrecioUnitario, "#,##0.00")                     'Contado
    Else
        Set itmX = lvVenta.ListItems.Add(, "N" & cCuota.ItemData(cCuota.ListIndex) & "P" & aIDPlan & "A" & aIdArticulo, cCuota.Text)
        itmX.SubItems(3) = Format(PrecioUnitario, "#,##0.00")                     'Contado
    End If
        
    itmX.SubItems(1) = Cantidad
    itmX.SubItems(2) = Trim(Articulo)

    itmX.SubItems(4) = IVAArticulo(aIdArticulo)
    
    Dim aValorCuota As Currency, aTotalFinanciado As Currency
    aUnitario = PrecioUnitario * aIDCoeficiente
    aValorCuota = Redondeo(aUnitario / CCur(cCuota.Tag), mMRound)
    
    aUnitario = Format(aValorCuota * CCur(cCuota.Tag), FormatoMonedaP)
    aTotalFinanciado = Format(aUnitario * Cantidad, FormatoMonedaP)
    
    itmX.SubItems(6) = Format(aValorCuota * Cantidad, FormatoMonedaP)         'Cuota
    
    'Ajusto el subtotal con lo que me da la cuota (SubTotal)
    itmX.SubItems(7) = Format(aTotalFinanciado, FormatoMonedaP) 'Total Financiado
    itmX.SubItems(8) = Format(aUnitario, FormatoMonedaP) 'Unitario Financiado
    
    TotalesSumo CCur(itmX.SubItems(7)), CCur(itmX.SubItems(4))
       
End Sub


Private Sub CargoDireccionesAuxiliares(aIdCliente As Long)

    On Error GoTo errCDA
    Dim rsDA As rdoResultset
    
    'Direcciones Auxiliares-----------------------------------------------------------------------
    Cons = "Select * from DireccionAuxiliar Where DAuCliente = " & aIdCliente & _
               " Order by DAuNombre"
    Set rsDA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsDA.EOF Then
        Do While Not rsDA.EOF
            cDireccion.AddItem Trim(rsDA!DAuNombre)
            cDireccion.ItemData(cDireccion.NewIndex) = rsDA!DAuDireccion
            If rsDA!DAuFactura Then gDirFactura = rsDA!DAuDireccion
            rsDA.MoveNext
        Loop
        
        If cDireccion.ListCount > 1 Then cDireccion.BackColor = Colores.Blanco
    End If
    rsDA.Close
    
    If Val(cDireccion.Tag) = 0 And cDireccion.ListCount > 0 And gDirFactura = 0 Then
        cDireccion.Text = cDireccion.List(0)
    Else
        If gDirFactura <> 0 Then BuscoCodigoEnCombo cDireccion, gDirFactura
    End If
    
    'If cDireccion.ListCount > 1 Then cDireccion.Visible = True Else cDireccion.Visible = False
    'cDireccion.Refresh
    
errCDA:
End Sub


Private Sub ListaDeSolicitudesPendientes(aIdCliente As Long)
    'Busca las solicitude no facturadas > a la ultima facturada.
    
    On Error GoTo errLista
    Screen.MousePointer = 11
    Dim rsLis As rdoResultset, aMaxSol As Long
    
    'Saco la Max Facturada----------------------------------------------------------------------------------------------------
    aMaxSol = 0
    Cons = "Select Max(SolCodigo) from Solicitud " & _
               " Where SolCliente = " & aIdCliente & " And SolProceso = " & TipoResolucionSolicitud.Facturada
    Set rsLis = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsLis.EOF Then If Not IsNull(rsLis(0)) Then aMaxSol = rsLis(0)
    rsLis.Close
    
    Dim bHayS As Boolean
    Cons = "Select * from Solicitud " & _
               " Where SolCliente = " & aIdCliente & " And SolCodigo > " & aMaxSol & _
               " And SolProceso <> " & TipoResolucionSolicitud.Facturada & " And SolFecha > '" & Format(DateAdd("m", -4, gFechaServidor), "mm/dd/yyyy") & "'"
    Set rsLis = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsLis.EOF Then bHayS = True Else bHayS = False
    rsLis.Close
    
    If Not bHayS Then Screen.MousePointer = 0: Exit Sub

    If MsgBox("El cliente tiene solicitudes pendientes no facturadas." & vbCrLf & "Para ver la lista de solicitudes presione 'SI'", vbQuestion + vbYesNo, "Hay Solicitudes No Facuradas") = vbNo Then Screen.MousePointer = 0: Exit Sub
    
    Cons = "Select SolCodigo + RSoArticulo + RSoTipoCuota, SolCodigo as 'Solicitud', SolFecha as 'Fecha', RSoCantidad as 'Q', ISNull(AEsNombre, ArtNombre) as 'Artículo', RSoValorEntrega as 'Entrega', RSoValorCuota as 'Valor Cuota'," & _
                        " SolComentarioS as 'Comentario Sol.', SolFResolucion as 'Resuelta', ResComentario as 'Comentario Res.' " & _
               " From Solicitud, RenglonSolicitud " & _
                    " LEFT OUTER JOIN ArticuloEspecifico ON RSoSolicitud = AEsDocumento And AEsTipoDocumento = 2 And RSoArticulo = AEsArticulo, " & _
                    " Articulo, SolicitudResolucion " & _
               " Where SolCliente = " & aIdCliente & " And SolCodigo > " & aMaxSol & _
               " And RSoArticulo = ArtID And SolCodigo = RSoSolicitud" & _
               " And SolProceso <> " & TipoResolucionSolicitud.Facturada & _
               " And SolCodigo = ResSolicitud " & _
               " And ResNumero = (Select MAX(ResNumero) From SolicitudResolucion Where SolCodigo = ResSolicitud)"
    
    Dim objLista As New clsListadeAyuda
    objLista.ActivarAyuda cBase, Cons, Me.Width - 100, 1, "Solicitudes Pendientes (No Facturadas)"
    Set objLista = Nothing
    
    '------------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 0
    Exit Sub
    
errLista:
    Screen.MousePointer = 0
End Sub


Private Function PrecioArticulo(lArticulo As Long, bConEntrega As Boolean, lMoneda As Long, lTCuota As Long, QCtas As Integer, _
                                      rUnitarioF As Currency, rCuotaF As Currency, rPlan As Long, rbNoHabPlan As Boolean, Optional rVariacionEspecifico As Currency) As Boolean
   
'19/8/2008 recibe la variación del artículo específico

'Si retorna true --> se puede vender, sino no hay coef p/calculo de ctas
    Dim miHayContado As Boolean
    PrecioArticulo = True
    rbNoHabPlan = False
    rUnitarioF = 0: rCuotaF = 0
    rPlan = paPlanPorDefecto
    
    
    If Not bConEntrega Then         'PROCESO PLAN SIN ENTREGA------------------------------------------------
        
        'Saco el valor de la cuota financiado
        Cons = "Select PViPrecio, PViHabilitado, PViPlan From PrecioVigente" _
                & " Where PVIArticulo = " & lArticulo _
                & " And PViMoneda = " & lMoneda _
                & " And PViTipoCuota = " & lTCuota _
                & " And PViHabilitado = 1"
        
        Dim bHayPrecio As Boolean
        bHayPrecio = False
        
        If rVariacionEspecifico = 0 Then
            Set RsArt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            bHayPrecio = Not RsArt.EOF
        Else
            bHayPrecio = False
        End If
        If bHayPrecio Then       'Hay Precios Grabados y están Habilitados
            rUnitarioF = Redondeo(RsArt!PViPrecio, mMRound)  'Format(RsArt!PViPrecio, "#,##0")                          'Precio de Unitario Financiaciado
            rCuotaF = Redondeo(RsArt!PViPrecio / QCtas, mMRound) 'Format(RsArt!PViPrecio / QCtas, "#,##0.00")     'Valor Cuota Finanaciado
            rPlan = RsArt!PViPlan
        
        Else            'No Hay Precios Grabados O no Están Habilitados
            Dim miPlan As Long, miCoef As Currency, miContado As Currency
            '1) Busco SI Hay precio Contado
            miHayContado = False: miContado = 0
            Cons = "Select PViPrecio, PViHabilitado, PViPlan From PrecioVigente" _
                    & " Where PVIArticulo = " & lArticulo _
                    & " And PViMoneda = " & lMoneda _
                    & " And PViTipoCuota = " & paTipoCuotaContado _
                    & " And PViHabilitado = 1"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux.EOF Then   'Si Hay contado busco el coeficiente p/tipo de Cuota y plan
                miHayContado = True
                miPlan = RsAux!PViPlan
                miContado = RsAux!PViPrecio
            End If
            RsAux.Close
                
            If Not miHayContado And lMoneda = paMonedaPesos Then
                'Si la moneda es pesos busco el precio Ctdo en U$S para hacer TC a pesos
                Cons = "Select PViPrecio, PViHabilitado, PViPlan From PrecioVigente" _
                        & " Where PVIArticulo = " & lArticulo _
                        & " And PViMoneda = " & paMonedaDolar _
                        & " And PViTipoCuota = " & paTipoCuotaContado _
                        & " And PViHabilitado = 1"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                If Not RsAux.EOF Then   'Si Hay contado busco el coeficiente p/tipo de Cuota y plan
                    miHayContado = True
                    miPlan = RsAux!PViPlan
                    miContado = RsAux!PViPrecio
                End If
                RsAux.Close
                
                If miHayContado Then miContado = Redondeo(miContado * paTCDolar, mMRound) 'Format(miContado * paTCDolar, "0")
            End If
                    
            If miHayContado Then    'Si hay ctdo busco coeficiente
                miContado = miContado + rVariacionEspecifico
                miCoef = 0
                'Busco el coeficiente p/tipo de Cuota y plan
                Cons = "Select * from Coeficiente" & _
                            " Where CoePlan = " & miPlan & _
                            " And CoeTipoCuota = " & lTCuota & _
                            " And CoeMoneda = " & lMoneda
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                If Not RsAux.EOF Then miCoef = RsAux!CoeCoeficiente
                RsAux.Close
                
                If miCoef <> 1 Then rbNoHabPlan = True
                                
                'If miCoef <> 1 Then Suceso
                If miCoef <> 0 Then
                    'rUnitarioF = Format(miContado * miCoef, "#,##0")                               'Precio de Unitario Financiaciado
                    'rUnitarioF = (Format(rUnitarioF / QCtas, "#,##0")) * QCtas                    'Ajusto el Valor de la Cta
                    'rCuotaF = Format(rUnitarioF / QCtas, "#,##0.00")                                'Valor Cuota Finanaciado
                    rUnitarioF = Redondeo(miContado * miCoef, mMRound)                               'Precio de Unitario Financiaciado
                    rUnitarioF = (Redondeo(rUnitarioF / QCtas, mMRound)) * QCtas                    'Ajusto el Valor de la Cta
                    rCuotaF = Redondeo(rUnitarioF / QCtas, mMRound)                                'Valor Cuota Finanaciado
                End If
                rPlan = miPlan
                                
            Else    'Si no hay contado dejo solicitar (Suceso)
                'Como no hay ni Contado, ni Credito pido el valor de la cuota
                rPlan = paPlanPorDefecto    'El plan no lo necesito         'No hay suceso.
            End If
            
        End If
        If rVariacionEspecifico = 0 Then RsArt.Close
    
    Else                                            'PROCESO PLAN CON ENTREGA------------------------------------------------
        miHayContado = False
        'Busco el precio contado del articulo para el plan seleccionado-------------------------
        Cons = "Select PViPrecio, PViHabilitado, PViPlan From PrecioVigente" _
                & " Where PVIArticulo = " & lArticulo _
                & " And PViMoneda = " & lMoneda _
                & " And PViTipoCuota = " & paTipoCuotaContado _
                & " And PViHabilitado = 1"
        Set RsArt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        
        'Si no hay CONTADO, hay que pedirlo para el plan que ingreso ---> voy a trabajar con el plan por defecto
        If Not RsArt.EOF Then
            rUnitarioF = Redondeo(RsArt!PViPrecio, mMRound)   'RsArt!PViPrecio    'Precio Unitario Contado
            rPlan = RsArt!PViPlan
            miHayContado = True
        Else
            rPlan = paPlanPorDefecto
        End If
        RsArt.Close
        
        If Not miHayContado And lMoneda = paMonedaPesos Then
            'Si la moneda es pesos busco el precio Ctdo en U$S para hacer TC a pesos
            Cons = "Select PViPrecio, PViHabilitado, PViPlan From PrecioVigente" _
                    & " Where PVIArticulo = " & lArticulo _
                    & " And PViMoneda = " & paMonedaDolar _
                    & " And PViTipoCuota = " & paTipoCuotaContado _
                    & " And PViHabilitado = 1"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux.EOF Then   'Si Hay contado busco el coeficiente p/tipo de Cuota y plan
                miHayContado = True
                rPlan = RsAux!PViPlan
                rUnitarioF = Redondeo(RsAux!PViPrecio, mMRound) 'RsArt!PViPrecio
            End If
            RsAux.Close
            
            If miHayContado Then rUnitarioF = Redondeo(rUnitarioF * paTCDolar, mMRound) 'Format(rUnitarioF * paTCDolar, "0")
        End If
        
        rUnitarioF = rUnitarioF + rVariacionEspecifico
        
        'Valido que exista un coeficiente para el calculo (TipoCuota, Plan, Moneda) Ya sea para el plan ingresado o el por defecto
        Cons = "Select * from Coeficiente" _
                & " Where CoePlan = " & rPlan _
                & " And CoeTipoCuota = " & lTCuota _
                & " And CoeMoneda = " & lMoneda
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If RsAux.EOF Then    'Si NO hay coeficientes NO SE VENDE
            PrecioArticulo = False
        End If
        RsAux.Close
    End If
    
End Function

Private Function ValidoRelacionTitularGarante() As Boolean
On Error GoTo errGral
    If Val(lGEdad.Tag) = 0 Then Exit Function
    
Dim aIdGarantia As Long, aIdTitular As Long
    
    ValidoRelacionTitularGarante = True
    
    If txtGarantia.Cliente.Codigo = 0 Or txtCliente.Cliente.Codigo = 0 Then Exit Function
    
    aIdGarantia = txtGarantia.Cliente.Codigo
    aIdTitular = txtCliente.Cliente.Codigo
    
    If aIdGarantia = aIdTitular Then
        MsgBox "La garantía no debe ser la misma persona que el titular de la solicitud.", vbExclamation, "Titular igual a Garantía"
        txtGarantia.Text = ""
        ValidoRelacionTitularGarante = False: Exit Function
    End If
    
    If txtCliente.Cliente.Tipo <> TC_Persona Then Exit Function
    If aIdTitular = gConyugeDelGarante Then Exit Function

Dim rsRel As rdoResultset, bHayRel As Boolean
    
    Screen.MousePointer = 11
    bHayRel = False
    
    Cons = "Select PReClienteDe from PersonaRelacion " & _
               " Where (PReClienteDe = " & aIdTitular & " And PReClienteEs = " & aIdGarantia & ")" & _
               " OR (PReClienteDe = " & aIdGarantia & " And PReClienteEs = " & aIdTitular & ")"
               
    Set rsRel = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsRel.EOF Then bHayRel = True
    rsRel.Close
    
    If bHayRel Then Screen.MousePointer = 0: Exit Function
    
    ValidoRelacionTitularGarante = False
     
    Dim aValor As Long, aInverso As Long
    Dim miLista As New clsListadeAyuda

listaRelacion:
    Cons = "Select -1 as RelCodigo, -1 as RelInvero, 0 as RelOrden, 'Cónyuge o Concubino' as 'Relación' " & _
                    " Union All " & _
                "Select RelCodigo, isnull(RelInverso, 0) as RelInvero, RelOrden, RelNombre as 'Relación' " & _
                " from Relaciones " & _
                " Order by RelOrden"
        
    aValor = miLista.ActivarAyuda(cBase, Cons, 4000, 3, "Qué es el Titular de la Garantía ?")
    If aValor <> 0 Then
        aValor = miLista.RetornoDatoSeleccionado(0)
        aInverso = miLista.RetornoDatoSeleccionado(1)
        Cons = miLista.RetornoDatoSeleccionado(3)
    End If
    Set miLista = Nothing
    
    If aValor = 0 Then
        If MsgBox("El ingreso de relaciones es obligatorio." & vbCrLf & "Quiere volver a var la lista de relaciones ?", vbQuestion + vbYesNo, "Las relaciones son obligatorioas") = vbNo Then
            txtGarantia.Text = ""
            Screen.MousePointer = 0
            Exit Function
        Else
            GoTo listaRelacion
        End If
    End If
    
    Dim bResult As Long
    bResult = MsgBox("Confirma que: " & vbCrLf & Trim(lblNombreCliente.Caption) & " es " & Trim(Cons) & " de " & Trim(lGarantia.Caption), vbQuestion + vbYesNo, "Confirma Grabar la Relación")
    
    Select Case bResult
        Case vbNo: GoTo listaRelacion
        
        Case vbYes:
                On Error GoTo errGrabar
                If aValor > -1 Then
                    ValidoRelacionTitularGarante = True
                
                    Cons = "Select * from PersonaRelacion " & _
                               " Where PReClienteDe = " & aIdTitular & " And PReClienteEs = " & aIdGarantia
                    Set rsRel = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    rsRel.AddNew
                    rsRel!PReClienteEs = aIdTitular
                    rsRel!PReRelacion = aValor
                    rsRel!PReClienteDe = aIdGarantia
                    rsRel.Update
                    
                    If aInverso <> 0 Then
                        rsRel.AddNew
                        rsRel!PReClienteEs = aIdGarantia
                        rsRel!PReRelacion = aInverso
                        rsRel!PReClienteDe = aIdTitular
                        rsRel.Update
                    End If
                    rsRel.Close
                
                Else    'Son conyuges o concuvinos
                    Dim bCancel As Boolean: bCancel = False
                    
                    Cons = "Select * from CPersona Where CPeCliente IN (" & aIdTitular & ", " & aIdGarantia & ")"
                    Set rsRel = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    Do While Not rsRel.EOF
                        If Not IsNull(rsRel!CPeConyuge) Then bCancel = True: Exit Do
                        rsRel.MoveNext
                    Loop
                    rsRel.Close
                    
                    If bCancel Then
                        MsgBox "Algunos de los clientes tiene asignado un cónyuge." & vbCrLf & _
                                    "Esta asignación debe realizarse manualmente en la ficha del cliente", vbExclamation, "Cónyuges Ya Asignados "
                        txtGarantia.Text = ""
                        ValidoRelacionTitularGarante = False
                        Screen.MousePointer = 0: Exit Function
                    Else
                        
                        Set rsRel = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                        Do While Not rsRel.EOF
                            
                            If IsNull(rsRel!CPeConyuge) Then
                                rsRel.Edit
                                If rsRel!CPeCliente = aIdTitular Then rsRel!CPeConyuge = aIdGarantia Else rsRel!CPeConyuge = aIdTitular
                                rsRel.Update
                            End If
                            
                            rsRel.MoveNext
                        Loop
                        rsRel.Close
                        ValidoRelacionTitularGarante = True
                        gConyugeDelGarante = aIdTitular
                        
                    End If
                    
                End If
    End Select
    Screen.MousePointer = 0
    Exit Function
    
errGrabar:
    clsGeneral.OcurrioError "Error al grabar las relaciones", Err.Description, "Grabar Relaciones"
    Screen.MousePointer = 0
    Exit Function

errGral:
    clsGeneral.OcurrioError "Error al procesar los datos de las relaciones", Err.Description, "Procesar Relaciones"
    Screen.MousePointer = 0
End Function

Private Function ValidoComboOK(mArticulo As Long, mPlan As Long) As Boolean
On Error GoTo errVCombo
Dim rsLoc As rdoResultset, rsM1 As rdoResultset
Dim bOK As Boolean

    ValidoComboOK = False
    'Saco todos los presupuestos en que está el artículo    ---------------------------------
    Cons = "Select PreID, PreArticulo, PreImporte from Presupuesto, PresupuestoArticulo  " & _
               " Where PreID = PArPresupuesto" & _
               " And PreMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) & _
               " And PArArticulo = " & mArticulo
    
    Set rsLoc = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not rsLoc.EOF
        
        bOK = True
        Cons = "Select * from PresupuestoArticulo" & _
                  " Where PArPresupuesto = " & rsLoc!PreID
                  
        Set rsM1 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsM1.EOF Then
            'Valido si esta la bonificaicion ----------------------------------------
            If Not IsNull(rsLoc!PreImporte) Then
                If rsLoc!PreImporte <> 0 Then
                    If Not Ingresado(rsLoc!PreArticulo, mPlan) Then bOK = False
                End If
            End If
            '------------------------------------------------------------------------
            If bOK Then
                Do While Not rsM1.EOF
                    If Not Ingresado(rsM1!PArArticulo, mPlan) Then bOK = False
                    If Not bOK Then Exit Do
                    rsM1.MoveNext
                Loop
            End If
        End If
        rsM1.Close
            
        If bOK Then Exit Do
        rsLoc.MoveNext
    Loop
    rsLoc.Close

    ValidoComboOK = bOK
    Exit Function

errVCombo:
    clsGeneral.OcurrioError "Error al validar artículo fuera de uso (caso combo).", Err.Description
End Function

Private Function code_ProcesoScript(mEvento As String, mIDCliente As Long, mIDSolicitud As Long) As String
On Error GoTo errCode
    
    If prmPlantillaPuente = 0 Then Exit Function
    Screen.MousePointer = 11
    code_ProcesoScript = ""
    
    Dim mFmt As Integer, mResult As String, mParams As String
    Dim objCode As New clsPlantillaI
    
    mFmt = 1
    
    mParams = "EVE=" & mEvento & "|" & _
                      "CLI=" & mIDCliente & "|" & _
                      "SOL=" & mIDSolicitud
                      
    If objCode.ProcesoPlantillaInteractiva(cBase, prmPlantillaPuente, mFmt, mResult, "", mParams, False) Then
        code_ProcesoScript = mResult
    End If
    
    Set objCode = Nothing
    Screen.MousePointer = 0
    Exit Function

errCode:
    clsGeneral.OcurrioError "Error al procesar código externo.", Err.Description
    Screen.MousePointer = 0
End Function

Private Function CargoTelefonos(a_IDCliente As Long)
On Error GoTo errFnc

    Cons = "Select IsNull(TTeOrden, 9999) as TTeOrden, * from Telefono, TipoTelefono" & _
                " Where TelCliente = " & a_IDCliente & _
                " And TelTipo = TTeCodigo " & _
                " Order by TTeOrden"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        cTelsT.AddItem Trim(RsAux!TTeNombre) & ": " & RsAux!TelNumero
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If cTelsT.ListCount > 0 Then
        lTelsT.Caption = lTelsT.Caption & " (" & cTelsT.ListCount & ")"
        cTelsT.ListIndex = 0
    End If
    
    Select Case cTelsT.ListCount
        Case Is < paQTelefonos
            cTelsT.BackColor = &HFF&
            cTelsT.ForeColor = vbWhite
        
        Case Else
            cTelsT.BackColor = lblNombreCliente.BackColor
            cTelsT.ForeColor = lblNombreCliente.ForeColor
    End Select
    
    Exit Function
errFnc:
End Function

Private Function zfn_InicializoControles()
On Error Resume Next

    ObtengoSeteoForm Me, (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Me.Height = 6140
    
    Me.BackColor = RGB(143, 188, 143) 'RGB(46, 139, 87)
    'Shape1.Shape = 0
    Shape1.BackColor = RGB(46, 139, 87) 'RGB(143, 188, 143)
    'Shape1.BorderStyle = 1
    Shape2.Shape = Shape1.Shape: Shape2.BackColor = Shape1.BackColor: Shape2.BorderStyle = Shape1.BorderStyle

    cEMailsT.BackColor = Shape2.BackColor
    lblNombreCliente.BackColor = RGB(220, 220, 220) 'RGB(60, 179, 113)
    labDireccion.BackColor = lblNombreCliente.BackColor
    lTEdad.BackColor = lblNombreCliente.BackColor: lRucCliente.BackColor = lblNombreCliente.BackColor
    lGEdad.BackColor = lblNombreCliente.BackColor: lGarantia.BackColor = lblNombreCliente.BackColor
    
    cTelsT.BackColor = lblNombreCliente.BackColor
    
    dis_CargoArrayMonedas
    
    FechaDelServidor
        
    cEMailsT.OpenControl cBase
    cEMailsT.IDUsuario = paCodigoDeUsuario
    
    'Cargo las monedas ------------------------------------------------------------------------------------
    Cons = "Select MonCodigo, MonSigno From Moneda Where MonFactura = 1 Order by MonSigno"
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        If dis_DisponibilidadPara(paCodigoDeSucursal, RsAux!MonCodigo) <> 0 Then
            cMoneda.AddItem Trim(RsAux!MonSigno)
            cMoneda.ItemData(cMoneda.NewIndex) = RsAux!MonCodigo
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If paMonedaFacturacion > 0 Then BuscoCodigoEnCombo cMoneda, paMonedaFacturacion
    '-----------------------------------------------------------------------------------------------------------
    
    'Cargo Comentarios de Solicitud ----------------------------------------------------------------------
    Cons = "Select ComCodigo, ComNombre From ComentarioSolicitud Order by ComNombre"
    CargoCombo Cons, cComentario, ""
    '-----------------------------------------------------------------------------------------------------------
    
    'Cargo las formas de pago-----------------------------------------------------------------------------
    cPago.AddItem cFPSEfectivo: cPago.ItemData(cPago.NewIndex) = TipoPagoSolicitud.Efectivo
    cPago.AddItem cFPSChequeD: cPago.ItemData(cPago.NewIndex) = TipoPagoSolicitud.ChequeDiferido
    '-----------------------------------------------------------------------------------------------------------
    
End Function

Private Function fnc_CallResolAutomatica(mID_SolAResolver As Long)
On Error GoTo errCRA
    
'    Dim oEjecAutoRes As New clsEjecutorAutoRes
'    oEjecAutoRes.AutoResolverCredito (mID_SolAResolver)
'    Set oEjecAutoRes = Nothing

    EjecutarApp App.Path & "\InvocoAutoResCred.exe", CStr(mID_SolAResolver)

    Exit Function
    
errCRA:
    clsGeneral.OcurrioError "Error al invocar la resolución automática.", Err.Description, "Resolución automática"
    
End Function

Private Sub txtCliente_BorroCliente()
    lbltitCliente.ForeColor = vbWhite
    lbltitCliente.FontBold = False
    LimpioDatosCliente
End Sub

Private Sub txtCliente_CambioTipoDocumento()
    lbltitCliente.ForeColor = vbWhite
    lbltitCliente.FontBold = False
    Select Case txtCliente.DocumentoCliente
        Case DC_CI
            lbltitCliente.Caption = "C.I.:"
        Case DC_RUT
            lbltitCliente.Caption = "R.U.T.:"
        Case Else
            If txtCliente.Cliente.TipoDocumento.Nombre = "" Then
                lbltitCliente.Caption = "Otro:"
            Else
                lbltitCliente.Caption = txtCliente.Cliente.TipoDocumento.Abreviacion
            End If
            lbltitCliente.ForeColor = &H80FF&
            lbltitCliente.FontBold = True
    End Select
End Sub

Private Sub txtCliente_Focus()
    Status.Panels(1).Text = "Ingrese el documento del cliente (F2=edita, F3=nuevo, F4=buscar)."
End Sub

Private Sub txtCliente_PresionoEnter()
    If txtCliente.Cliente.Codigo > 0 Then Foco cCuota
End Sub

Private Sub txtCliente_SeleccionoCliente()
On Error GoTo errSC
    LimpioDatosCliente
    fnc_CargoDatosCliente False
    txtCliente.BuscoComentariosAlerta txtCliente.Cliente.Codigo, True
    If txtCliente.Cliente.Tipo = TC_Persona Then ValidoMayorDeEdad True, False
    code_ProcesoScript "SOL01", txtCliente.Cliente.Codigo, 0
    If cDireccion.ListCount > 1 Then
        cDireccion.SetFocus
    Else
        If cEMailsT.GetDirecciones = "" Then cEMailsT.SetFocus Else txtGarantia.SetFocus
    End If
    Exit Sub
errSC:
    clsGeneral.OcurrioError "Error al cargar la ficha del cliente.", Err.Description
    txtCliente.Text = ""
End Sub

Private Sub txtGarantia_BorroCliente()
    lGarantia.Caption = ""
    lGEdad.Caption = ""
    Shape1.BackColor = RGB(46, 139, 87)
End Sub

Private Sub txtGarantia_CambioTipoDocumento()
    lblInfoAval.ForeColor = vbWhite
    lblInfoAval.FontBold = False
    If txtGarantia.DocumentoCliente = DC_Otros Then
        lblInfoAval.ForeColor = &H80FF&
        lblInfoAval.FontBold = True
    End If
End Sub

Private Sub txtGarantia_Focus()
    Status.Panels(1).Text = "Ingrese la cédula de identidad de la garantía."
End Sub

Private Sub txtGarantia_PresionoEnter()
    cCuota.SetFocus
End Sub

Private Sub txtGarantia_SeleccionoCliente()
    fnc_CargoDatosGarantia
    txtGarantia.BuscoComentariosAlerta txtGarantia.Cliente.Codigo, True
    ValidoMayorDeEdad False, True
    ValidoRelacionTitularGarante
    If txtGarantia.Cliente.Codigo > 0 Then Foco cCuota
End Sub

Private Function CalculoEdad(ByVal fechaNacimiento As Date) As Integer

    If DateAdd("yyyy", DateDiff("yyyy", fechaNacimiento, Date), fechaNacimiento) > Date Then
        CalculoEdad = DateDiff("yyyy", fechaNacimiento, Date) - 1
    Else
        CalculoEdad = DateDiff("yyyy", fechaNacimiento, Date)
    End If

End Function

Private Function CargoObjetoArticuloDeColeccion(ByVal IDArticulo As Long) As clsArticulo
On Error GoTo errCOA
'    Screen.MousePointer = 11
'    Set oArtEdicion = New clsArticulo
'    Cons = "SELECT ArtNombre, ArtID, IsNull(ArtEnVentaXMayor, 1) VtaXMayor From Articulo WHERE ArtID = " & idArticulo
'    Dim rsA As rdoResultset
'    Set rsA = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
'    If Not rsA.EOF Then
'        oArtEdicion.ID = rsA("ArtID")
'        oArtEdicion.NombreArticulo = Trim(rsA("ArtNombre"))
'        oArtEdicion.VentaXMayor = rsA("VtaXMayor")
'    End If
'    rsA.Close
    Dim iIx As Integer
    Dim oArt As clsArticulo
    For Each oArt In colArtsGrilla
        If oArt.ID = IDArticulo Then
            Set CargoObjetoArticuloDeColeccion = oArt
            Exit Function
        End If
    Next
    Screen.MousePointer = 0
    Exit Function
errCOA:
    clsGeneral.OcurrioError "Error al cargar la información del artículo.", Err.Description
    Screen.MousePointer = 0
End Function
