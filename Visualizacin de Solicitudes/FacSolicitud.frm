VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FacSolicitud 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitudes de Compra"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   615
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
   Icon            =   "FacSolicitud.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   9465
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox tVendedor 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   5380
      MaxLength       =   2
      TabIndex        =   23
      Top             =   4635
      Width           =   375
   End
   Begin VB.TextBox tInformacion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   0
      Width           =   9472
   End
   Begin VB.ComboBox cComentario 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1080
      TabIndex        =   27
      Top             =   4995
      Width           =   5775
   End
   Begin VB.ComboBox cPago 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   2760
      TabIndex        =   21
      Top             =   4635
      Width           =   1695
   End
   Begin VB.TextBox tEntregaT 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1080
      MaxLength       =   12
      TabIndex        =   19
      Top             =   4635
      Width           =   1095
   End
   Begin VB.TextBox tEntrega 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   6960
      MaxLength       =   12
      TabIndex        =   15
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ComboBox cCuota 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox tUsuario 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   6480
      MaxLength       =   3
      TabIndex        =   25
      Top             =   4635
      Width           =   375
   End
   Begin VB.ComboBox cArticulo 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1320
      Style           =   1  'Simple Combo
      TabIndex        =   9
      Top             =   2640
      Width           =   3855
   End
   Begin VB.TextBox tCantidad 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   5160
      MaxLength       =   5
      TabIndex        =   11
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox tUnitario 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   5760
      MaxLength       =   12
      TabIndex        =   13
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ComboBox cMoneda 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   8520
      TabIndex        =   29
      Text            =   "cMoneda"
      Top             =   840
      Width           =   855
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   30
      Top             =   5430
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   11829
            MinWidth        =   2
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   1799
            MinWidth        =   1482
            Text            =   "F2-Modificar "
            TextSave        =   "F2-Modificar "
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   1482
            MinWidth        =   1482
            Text            =   "F3-Nuevo"
            TextSave        =   "F3-Nuevo"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   1482
            MinWidth        =   1482
            Text            =   "F4-Buscar"
            TextSave        =   "F4-Buscar"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox tCi 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   327681
      ForeColor       =   12582912
      MaxLength       =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#.###.###-#"
      PromptChar      =   "_"
   End
   Begin ComctlLib.ListView lvVenta 
      Height          =   1575
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   9
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Plan"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Cant."
         Object.Width           =   617
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Artículo"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Contado x1"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "I.V.A."
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Entrega"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Cuota"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Sub Total"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   8
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Financiado x1"
         Object.Width           =   0
      EndProperty
   End
   Begin MSMask.MaskEdBox tGarantia 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   327681
      ForeColor       =   12582912
      MaxLength       =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#.###.###-#"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox tRuc 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   327681
      ForeColor       =   12582912
      PromptInclude   =   0   'False
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99 999 999 9999"
      PromptChar      =   "_"
   End
   Begin VB.Label lUnitario 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cuota x&1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   2415
      Width           =   1215
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   960
      Left            =   240
      Picture         =   "FacSolicitud.frx":0442
      Top             =   320
      Width           =   3795
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "&Vendedor:"
      Height          =   255
      Left            =   4560
      TabIndex        =   22
      Top             =   4635
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
      TabIndex        =   16
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lTEdad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "90"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8925
      TabIndex        =   54
      Top             =   1320
      UseMnemonic     =   0   'False
      Width           =   315
   End
   Begin VB.Label lGEdad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "90"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8925
      TabIndex        =   53
      Top             =   2040
      UseMnemonic     =   0   'False
      Width           =   315
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Edad:"
      Height          =   255
      Left            =   8475
      TabIndex        =   52
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Edad:"
      Height          =   255
      Left            =   8475
      TabIndex        =   51
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Pag&o:"
      Height          =   255
      Left            =   2280
      TabIndex        =   20
      Top             =   4635
      Width           =   495
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Co&mentarios:"
      Height          =   255
      Left            =   105
      TabIndex        =   26
      Top             =   4995
      Width           =   975
   End
   Begin VB.Label lGarantia 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rodriguez Fernandez, Rodrigo Bernardino"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   49
      Top             =   2040
      UseMnemonic     =   0   'False
      Width           =   4695
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   3000
      TabIndex        =   48
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "C.I. &Garantía:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "E&ntrega:"
      Height          =   255
      Left            =   105
      TabIndex        =   18
      Top             =   4635
      Width           =   735
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Entrega"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6960
      TabIndex        =   14
      Top             =   2415
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   9480
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label lSubTotalF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   8160
      TabIndex        =   47
      Top             =   2640
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
      TabIndex        =   46
      Top             =   2415
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
      Height          =   255
      Left            =   5820
      TabIndex        =   24
      Top             =   4635
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
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6840
      TabIndex        =   45
      Top             =   5040
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
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6840
      TabIndex        =   44
      Top             =   4800
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
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6960
      TabIndex        =   43
      Top             =   4560
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
      TabIndex        =   42
      Top             =   4560
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
      TabIndex        =   41
      Top             =   4800
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
      TabIndex        =   40
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Artículo"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   2415
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
      TabIndex        =   10
      Top             =   2415
      Width           =   615
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Moneda"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8520
      TabIndex        =   28
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&C.I. Cliente:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&R.U.C.:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   3000
      TabIndex        =   39
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label labNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rodriguez Fernandez, Rodrigo Bernardino"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   38
      Top             =   1320
      UseMnemonic     =   0   'False
      Width           =   4695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      Height          =   255
      Left            =   3000
      TabIndex        =   37
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label labDireccion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Niagara 2345"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   36
      Top             =   1680
      UseMnemonic     =   0   'False
      Width           =   5535
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Solicitud de Compra"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6480
      TabIndex        =   35
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Documento"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   34
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label labFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10-Dic-1998"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5160
      TabIndex        =   33
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5160
      TabIndex        =   32
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "R.U.C. 21.025996.0012"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   31
      Top             =   360
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   1160
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1220
      Width           =   9255
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "       &Plan"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   9255
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuEmitir 
         Caption         =   "&Grabar Solicitud"
         Enabled         =   0   'False
         Shortcut        =   ^G
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
   Begin VB.Menu MnuMoussePersona 
      Caption         =   "&MoussePersona"
      Visible         =   0   'False
      Begin VB.Menu MnuMoCliente 
         Caption         =   "Menú Cliente"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuLineaMP1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuNuevoCliente 
         Caption         =   "&Nuevo Cliente                 F3"
      End
      Begin VB.Menu MnuPBuscar 
         Caption         =   "&Buscar Clientes               F4"
      End
      Begin VB.Menu MnuFichaCliente 
         Caption         =   "Ir a &ficha de Cliente         F2"
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
      Begin VB.Menu MnuLineaMP2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCancelarMP 
         Caption         =   "&Cancelar"
      End
   End
   Begin VB.Menu MnuMousseEmpresa 
      Caption         =   "MousseEmpresa"
      Visible         =   0   'False
      Begin VB.Menu MnuMoEmpresa 
         Caption         =   "Menú Empresa"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuLineaME1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuNuevaEmpresa 
         Caption         =   "&Nueva Empresa                  F2"
      End
      Begin VB.Menu MnuEBuscar 
         Caption         =   "&Buscar Empresas                F4"
      End
      Begin VB.Menu MnuFichaEmpresa 
         Caption         =   "Ir a &Ficha de Empresa         F3"
      End
      Begin VB.Menu MnuLineaME2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCancelarME 
         Caption         =   "&Cancelar"
      End
   End
   Begin VB.Menu MnuMousseGarantia 
      Caption         =   "MousseGarantia"
      Visible         =   0   'False
      Begin VB.Menu MnuMoGarantia 
         Caption         =   "Menú Garantía"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuLineaMG1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGNuevo 
         Caption         =   "&Nuevo Cliente           F2"
      End
      Begin VB.Menu MnuGBuscar 
         Caption         =   "&Buscar Clientes         F4"
      End
      Begin VB.Menu MnuGFicha 
         Caption         =   "Ir a &Ficha de Cliente"
      End
      Begin VB.Menu MnuLineaMG2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCancelarMG 
         Caption         =   "&Cancelar"
      End
   End
End
Attribute VB_Name = "FacSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sConEntrega As Boolean
Dim sDistribuir As Boolean          'Para redistribuir las entregas

'Forms.-----------------------------------------
Private frmPersona As MaCPersona
Private frmEmpresa As MaCEmpresa

'RDO.------------------------------------------
Private RsBoleta As rdoResultset
Private RsArt As rdoResultset


Private Sub cArticulo_GotFocus()

    cArticulo.SelStart = 0
    cArticulo.SelLength = Len(cArticulo.Text)
    Status.Panels(1).Text = "Ingrese el código o nombre del artículo."
    
End Sub

Private Sub cArticulo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(cArticulo.Text) <> vbNullString Then
            If Not ValidoBuscarArticulo Then Exit Sub
            
            If Not IsNumeric(cArticulo.Text) Then
                BuscoArticuloXNombre
            Else
                BuscoArticuloxCodigo CLng(cArticulo.Text)
            End If
            
            If cArticulo.ListIndex > -1 Then Foco tCantidad
            
        Else
            If tEntregaT.Enabled Then tEntregaT.SetFocus Else cPago.SetFocus
        End If
    End If

End Sub

Private Sub cComentario_Change()
    Selecciono cComentario, cComentario.Text, gTecla
End Sub

Private Sub cComentario_KeyDown(KeyCode As Integer, Shift As Integer)
    gTecla = KeyCode
    gIndice = cComentario.ListIndex
End Sub

Private Sub cComentario_KeyUp(KeyCode As Integer, Shift As Integer)
    ComboKeyUp cComentario
End Sub

Private Sub cComentario_LostFocus()
    gIndice = -1
    cComentario.SelLength = 0
End Sub

Private Sub cCuota_Change()
    Selecciono cCuota, cCuota.Text, gTecla
    LimpioRenglon
    If cCuota.ListIndex <> -1 Then Call cCuota_Click
End Sub

Private Sub cCuota_Click()

    On Error GoTo errBuscar
    LimpioRenglon
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
    Screen.MousePointer = 0
    clsError.MuestroError "Ocurrió un error al buscar el tipo de cuota."
    cCuota.Text = ""
End Sub

Private Sub cCuota_GotFocus()

    cCuota.SelStart = 0
    cCuota.SelLength = Len(cCuota.Text)
    
    Status.Panels(1).Text = "Seleccione la financiación del artículo."
    
End Sub

Private Sub cCuota_KeyDown(KeyCode As Integer, Shift As Integer)

    gTecla = KeyCode
    gIndice = cCuota.ListIndex
    
End Sub

Private Sub cCuota_KeyPress(KeyAscii As Integer)

    cCuota.ListIndex = gIndice
    If KeyAscii = vbKeyReturn Then
        If Trim(cCuota.Text) = "" Then
            If tEntregaT.Enabled Then
                tEntregaT.SetFocus
            Else
                cPago.SetFocus
            End If
        Else
            cArticulo.SetFocus
        End If
    End If

End Sub

Private Sub cCuota_KeyUp(KeyCode As Integer, Shift As Integer)

    ComboKeyUp cCuota
    
End Sub

Private Sub cCuota_LostFocus()

    gIndice = -1
    cCuota.SelLength = 0
    
End Sub

Private Sub cMoneda_Change()

    Selecciono cMoneda, cMoneda.Text, gTecla

End Sub

Private Sub cMoneda_Click()
    LimpioRenglon
End Sub

Private Sub cMoneda_GotFocus()

    cMoneda.SelStart = 0
    cMoneda.SelLength = Len(cMoneda.Text)
    Status.Panels(1).Text = "Seleccione una moneda para facturar la solicitud."

End Sub

Private Sub cMoneda_KeyDown(KeyCode As Integer, Shift As Integer)

    gTecla = KeyCode
    gIndice = cMoneda.ListIndex

End Sub

Private Sub cMoneda_KeyPress(KeyAscii As Integer)

    cMoneda.ListIndex = gIndice
    If KeyAscii = vbKeyReturn Then tCi.SetFocus

End Sub

Private Sub cMoneda_KeyUp(KeyCode As Integer, Shift As Integer)

    ComboKeyUp cMoneda

End Sub

Private Sub cMoneda_LostFocus()

    gIndice = -1
    cMoneda.SelLength = 0
    
End Sub

Private Sub cPago_Change()
    Selecciono cPago, cPago.Text, gTecla
End Sub

Private Sub cPago_GotFocus()
    cPago.SelStart = 0
    cPago.SelLength = Len(cPago.Text)
    
    Status.Panels(1).Text = "Seleccione la forma de pago de la solicitud."
End Sub

Private Sub cPago_KeyDown(KeyCode As Integer, Shift As Integer)
    gTecla = KeyCode
    gIndice = cPago.ListIndex
End Sub

Private Sub cPago_KeyPress(KeyAscii As Integer)

    cPago.ListIndex = gIndice
    If KeyAscii = vbKeyReturn Then Foco tVendedor
    
End Sub

Private Sub cPago_KeyUp(KeyCode As Integer, Shift As Integer)
    ComboKeyUp cPago
End Sub

Private Sub cPago_LostFocus()

    gIndice = -1
    cPago.SelLength = 0
    
End Sub

Private Sub Form_Activate()

    Screen.MousePointer = vbDefault
    DoEvents

    Cons = CargoInformacionDia
    tInformacion.Text = Cons
       
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF2            'Modificar Ficha
            If Mid(labNombre.Tag, 1, 1) = "E" Then IrAEmpresa       'Empresa
            
            If Mid(labNombre.Tag, 1, 1) = "P" Then IrACliente True, False    'Persona
    End Select
    
End Sub

Private Sub Form_Load()
    
    'Cargo las monedas ------------------------------------------------------------------------------------
    Cons = "Select MonCodigo, MonSigno From Moneda Order by MonSigno"
    CargoCombo Cons, cMoneda, ""
    If paMonedaFacturacion > 0 Then BuscoCodigoEnCombo cMoneda, paMonedaFacturacion
    '-----------------------------------------------------------------------------------------------------------
    
    'Cargo Comentarios de Solicitud -----------------------------------------------------------------------
    Cons = "Select ComCodigo, ComNombre From ComentarioSolicitud Order by ComNombre"
    CargoCombo Cons, cComentario, ""
    '-----------------------------------------------------------------------------------------------------------
    
    'Cargo las formas de pago-----------------------------------------------------------------------------
    cPago.AddItem cFPSEfectivo
    cPago.ItemData(cPago.NewIndex) = TipoPagoSolicitud.Efectivo
    cPago.AddItem cFPSChequeD
    cPago.ItemData(cPago.NewIndex) = TipoPagoSolicitud.ChequeDiferido
    '-----------------------------------------------------------------------------------------------------------
    
    LimpioTodaLaFicha False
    SetearLView lvValores.Grilla Or lvValores.FullRow, lvVenta
    labFecha.Caption = Format(Date, "d-Mmm-yyyy")

    gIndice = -1
        
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.Panels(1).Text = vbNullString
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Forms(Forms.Count - 2).SetFocus

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
    Foco tGarantia
End Sub

Private Sub Label17_Click()
    Foco cCuota
End Sub

Private Sub Label2_Click()

    tRuc.SelStart = 0
    tRuc.SelLength = 15
    tRuc.SetFocus

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

Private Sub Label5_Click()
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
                BuscoCodigoEnCombo cCuota, Mid(lvVenta.SelectedItem.Key, 2, InStr(lvVenta.SelectedItem.Key, "P") - 2)
                'Saco el Codigo del Articulo
                cArticulo.AddItem Trim(lvVenta.SelectedItem.SubItems(2))
                
                cArticulo.ItemData(cArticulo.NewIndex) = Mid(lvVenta.SelectedItem.Key, InStr(lvVenta.SelectedItem.Key, "A") + 1, Len(lvVenta.SelectedItem.Key))
                cArticulo.ListIndex = 0
                
                tCantidad.Text = lvVenta.SelectedItem.SubItems(1)
                tUnitario.Tag = lvVenta.SelectedItem.SubItems(3)        'P.U. Ctdo
                'Cuota
                If Not sConEntrega Then
                    tUnitario.Text = Format(CCur(lvVenta.SelectedItem.SubItems(6)) / CCur(tCantidad.Text), FormatoMonedaP)
                Else
                    tUnitario.Text = tUnitario.Tag
                End If
                
                lSubTotalF.Caption = lvVenta.SelectedItem.SubItems(7)   'Subtotal
                lSubTotalF.Tag = lvVenta.SelectedItem.SubItems(8)        'P.U. Financiado
                
                If sConEntrega Then
                    If Trim(lSubTotalF.Caption) = "" Then       'VALIDO QUE SE HAYA REALIZADO LA DISTRIBUCION
                        MsgBox "Aún no ha ingresado el valor de entrega total de la solicitud o no se realizó la distribución automática.", vbExclamation, "ATENCIÓN"
                        LimpioRenglon
                        If tEntregaT.Enabled Then tEntregaT.SetFocus
                    Else
                        tEntrega.Text = lvVenta.SelectedItem.SubItems(5)
                        IngresoDeEntrega True
                        Foco tEntrega
                    End If
                Else
                    TotalesResto CCur(lvVenta.SelectedItem.SubItems(7)), CCur(lvVenta.SelectedItem.SubItems(4))
                    lvVenta.ListItems.Remove lvVenta.SelectedItem.Index
                    Foco tCantidad
                End If
                
                If lvVenta.ListItems.Count = 0 Then
                    cMoneda.Enabled = True
                    MnuEmitir.Enabled = False
                End If
                
            Case vbKeyDelete    'ELIMINO EL RENGLON----------------------------------------------------------------------------------
                
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
                    cMoneda.Enabled = False
                    MnuEmitir.Enabled = True
                Else
                    cMoneda.Enabled = True
                    MnuEmitir.Enabled = False
                End If
            
            Case vbKeyReturn
                If tEntregaT.Enabled Then
                    tEntregaT.SetFocus
                Else
                    cPago.SetFocus
                End If
            
        End Select
    End If
    Exit Sub

ErrlvKD:
    clsError.MuestroError "Ocurrio un error inesperado."
End Sub

Private Sub MnuEBuscar_Click()

    BuscarClientes TipoCliente.Empresa
    
End Sub

Private Sub MnuEmitir_Click()

    AccionGrabar

End Sub

Private Sub MnuFichaCliente_Click()
    
    IrACliente True, False
    
End Sub

Private Sub MnuFichaEmpresa_Click()
    IrAEmpresa
End Sub

Private Sub MnuGBuscar_Click()
    BuscarGarantia
End Sub

Private Sub BuscarGarantia()
    
    Screen.MousePointer = 11
    BuscarCliente.pSeleccionado = 0
    BuscarCliente.pSeleccionadoTipo = TipoCliente.Persona
    BuscarCliente.Show vbModal, Me
    Me.Refresh
    
    On Error GoTo errCargar
    If BuscarCliente.pSeleccionado <> 0 And BuscarCliente.pSeleccionadoTipo = TipoCliente.Persona Then
        tGarantia.Text = FormatoCedula
        lGarantia.Caption = ""
        tGarantia.Tag = ""
        'Cargo los datos del Cliente recien ingresado---------------------------------------------
        Cons = "Select CliCodigo, CliCIRuc, CliCategoria, CliDireccion, CPersona.* From Cliente, CPersona " _
                & " Where CliCodigo = " & BuscarCliente.pSeleccionado _
                & " And CliTipo = " & TipoCliente.Persona _
                & " And CliCodigo = CPeCliente"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        CargoDatosGarantia
        BuscoComentariosAlerta RsAux!CliCodigo, lGarantia.Caption, Alerta:=True
        RsAux.Close
        '--------------------------------------------------------------------------------------------------
        
        ValidoMayorDeEdad False, True
        If Trim(lGarantia.Caption) <> "" Then cCuota.SetFocus
        
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    clsError.MuestroError "Ocurrió un error al cargar los datos del cliente."
End Sub

Private Sub MnuGFicha_Click()

    IrACliente False, True
    
End Sub

Private Sub MnuGNuevo_Click()

    IrAClienteNuevo False, True
    
End Sub

Private Sub MnuLimpiar_Click()

   LimpioTodaLaFicha True
    
End Sub

Private Sub LimpioTodaLaFicha(FocoCI As Boolean, Optional DejarGarantia As Boolean = False)
    
    LimpioRenglon
    lvVenta.ListItems.Clear
    cMoneda.Enabled = True
    
    labNombre.Caption = ""
    labDireccion.Caption = ""
    labNombre.Tag = vbNullString
    labDireccion.Tag = ""
    lTEdad.Caption = ""
    
    tRuc.Text = vbNullString
    tUsuario.Text = vbNullString
    tUsuario.Tag = vbNullString
    tVendedor.Text = ""
    
    If Not DejarGarantia Then
        tGarantia.Text = FormatoCedula
        tGarantia.Tag = ""
        lGarantia.Caption = ""
        lGEdad.Caption = ""
    End If
    
    cCuota.Text = ""
    tEntregaT.Text = ""
    cComentario.Text = ""
    cPago.ListIndex = 0
    
    labTotal.Caption = "0.00"
    labIva.Caption = "0.00"
    labSubTotal.Caption = "0.00"

    HabilitoEntrega
    IngresoDeEntrega False
    
    If FocoCI Then tCi.SetFocus
    
End Sub


Private Sub MnuNuevaEmpresa_Click()
    IrAEmpresaNuevo
End Sub

Private Sub MnuNuevoCliente_Click()
    IrAClienteNuevo True, False
End Sub

Private Sub MnuPBuscar_Click()
    BuscarClientes TipoCliente.Persona
End Sub

Private Sub BuscarClientes(aTipoCliente As Integer)
    
    Screen.MousePointer = 11
    BuscarCliente.pSeleccionado = 0
    BuscarCliente.pSeleccionadoTipo = aTipoCliente
    BuscarCliente.Show vbModal, Me
    Me.Refresh
    
    On Error GoTo errCargar
    If BuscarCliente.pSeleccionado <> 0 Then
        LimpioDatosCliente
        If BuscarCliente.pSeleccionadoTipo = TipoCliente.Persona Then   'PERSONA---------------
        
            Cons = "Select CliCodigo, CliCIRuc, CliCategoria, CliDireccion, CPersona.* From Cliente, CPersona " _
                    & " Where CliCodigo = " & BuscarCliente.pSeleccionado _
                    & " And CliTipo = " & TipoCliente.Persona _
                    & " And CliCodigo = CPeCliente"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            CargoDatosPersona
            BuscoComentariosAlerta RsAux!CliCodigo, labNombre.Caption, Alerta:=True
            RsAux.Close
            
            ValidoMayorDeEdad True, False
            
        Else                                                                                        'EMPRESA----------------
            Cons = "Select CliCodigo, CliCIRuc, CliCategoria, CEmFantasia, CEmNombre, CliDireccion from Cliente, CEmpresa" _
                    & " Where CliCodigo = " & BuscarCliente.pSeleccionado _
                    & " And CliTipo = " & TipoCliente.Empresa _
                    & " And CliCodigo = CEmCliente"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            CargoDatosEmpresa
            BuscoComentariosAlerta RsAux!CliCodigo, labNombre.Caption, Alerta:=True
            RsAux.Close
        End If
        
        If Trim(labNombre.Caption) <> "" Then tGarantia.SetFocus
        
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    clsError.MuestroError "Ocurrió un error al cargar los datos del cliente."
End Sub

Private Sub MnuPEmpleo_Click()

    If Mid(labNombre.Tag, 1, 1) <> "P" Then Exit Sub
    IrAEmpleo CLng(Mid(labNombre.Tag, 2, Len(labNombre.Tag) - 1))
    
End Sub

Private Sub IrAEmpleo(Cliente As Long)

Dim fEmpleo As New MaEmpleo
    
    Screen.MousePointer = 11
    fEmpleo.pCiCliente = ""
    fEmpleo.pCodCliente = Cliente
    fEmpleo.pNomCliente = Trim(labNombre.Caption)
    fEmpleo.Show vbModal, Me
    DoEvents
    Set fEmpleo = Nothing
    
End Sub
Private Sub IrAReferencia(Cliente As Long)

Dim fReferencia As New MaAsigReferencia
    
    Screen.MousePointer = 11
    fReferencia.pCiCliente = ""
    fReferencia.pCodCliente = Cliente
    fReferencia.pNomCliente = Trim(labNombre.Caption)
    fReferencia.Show vbModal, Me
    DoEvents
    Set fReferencia = Nothing
    
End Sub

Private Sub IrATitulo(Cliente As Long)

Dim fTitulo As New MaTitulo
    
    Screen.MousePointer = 11
    fTitulo.pCiCliente = ""
    fTitulo.pCodCliente = Cliente
    fTitulo.pNomCliente = Trim(labNombre.Caption)
    fTitulo.Show vbModal, Me
    DoEvents
    Set fTitulo = Nothing
    
End Sub
Private Sub MnuPReferencia_Click()
    If Mid(labNombre.Tag, 1, 1) <> "P" Then Exit Sub
    IrAReferencia CLng(Mid(labNombre.Tag, 2, Len(labNombre.Tag) - 1))
End Sub

Private Sub MnuPTitulo_Click()
    If Mid(labNombre.Tag, 1, 1) <> "P" Then Exit Sub
    IrATitulo CLng(Mid(labNombre.Tag, 2, Len(labNombre.Tag) - 1))
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
            If Val(tCantidad.Text) > 0 Then tUnitario.SetFocus
        End If
    End If

End Sub

Private Sub tCantidad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.Panels(1).Text = "Ingrese la cantidad de artículos."
End Sub

Private Sub tCi_GotFocus()

    tCi.SelStart = 0
    tCi.SelLength = Len(tCi.Text)
    
    Status.Panels(1).Text = "Ingrese la cédula de identidad del cliente."
    
End Sub

Private Sub tCi_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF4: If Shift = 0 Then BuscarClientes TipoCliente.Persona        'Buscar Clientes
        Case vbKeyF3: IrAClienteNuevo True, False                  'Nuevo Cliente
        
        Case 93             'Boton Derecho
            If Mid(labNombre.Tag, 1, 1) = "P" Then
                MnuPEmpleo.Enabled = True
                MnuPReferencia.Enabled = True
                MnuPTitulo.Enabled = True
                MnuFichaCliente.Enabled = True
            Else
                MnuPEmpleo.Enabled = False
                MnuPReferencia.Enabled = False
                MnuPTitulo.Enabled = False
                MnuFichaCliente.Enabled = False
            End If
            PopupMenu MnuMoussePersona, , tCi.Left + (tCi.Width / 2), (tCi.Top + tCi.Height) - (tRuc.Height / 2)
        
    End Select

End Sub

Private Sub TCI_KeyPress(KeyAscii As Integer)

    On Error GoTo ErrTCK

    If KeyAscii = vbKeyReturn Then
        Dim aCi As String
        Screen.MousePointer = 11
        LimpioTodaLaFicha False
        'Valido la Cédula ingresada----------
        If Trim(tCi.Text) <> FormatoCedula Then
            aCi = QuitoFormatoCedula(tCi.Text)
            If Len(aCi) <> 8 Or Not CedulaValida(aCi) Then
                Screen.MousePointer = vbDefault
                MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
        End If
        'Busco el Cliente -----------------------
        If Trim(tCi.Text) <> FormatoCedula Then
            BuscoClienteCI QuitoFormatoCedula(tCi.Text)
            If labNombre.Tag <> vbNullString Then tGarantia.SetFocus
        Else
            tRuc.SetFocus
        End If
        Screen.MousePointer = 0
    End If
    Exit Sub
    
ErrTCK:
    Screen.MousePointer = 0
    clsError.MuestroError "Ocurrio un error al buscar la cédula de identidad."
End Sub

Private Sub BuscoClienteCI(Codigo As String)
On Error GoTo errBuscar

    tRuc.Text = ""
    If Codigo = "" Then Exit Sub
    Screen.MousePointer = vbHourglass

    Cons = "Select CliCodigo, CliCIRuc, CliCategoria, CliDireccion, CPersona.* From Cliente, CPersona " _
            & " Where CliCiRuc = '" & Codigo & "'" _
            & " And CliTipo = " & TipoCliente.Persona _
            & " And CliCodigo = CPeCliente"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)

    If RsAux.EOF Then
        Screen.MousePointer = 0
        RsAux.Close
        If MsgBox("No existe un cliente para la cédula de indentidad ingresada. Desea ingresarlo.", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
                    
            Set frmPersona = New MaCPersona
            frmPersona.pTipoLlamado = TipoLlamado.CreditoAClientes
            frmPersona.pClienteSeleccionado = 0
            frmPersona.pConyugeDe = 0
            frmPersona.tCi.Text = tCi.Text
            frmPersona.Show vbModal, Me
            
            If frmPersona.pClienteSeleccionado > 0 Then
                'Cargo los datos del Cliente recien ingresado---------------------------------------------
                Cons = "Select CliCodigo,CliCIRuc, CliCategoria, CliDireccion, CPersona.* From Cliente, CPersona " _
                    & " Where CliCodigo = " & frmPersona.pClienteSeleccionado _
                    & " And CliTipo = " & TipoCliente.Persona _
                    & " And CliCodigo = CPeCliente"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                LimpioDatosCliente
                CargoDatosPersona
                BuscoComentariosAlerta RsAux!CliCodigo, labNombre.Caption, Alerta:=True
                RsAux.Close
                ValidoMayorDeEdad True, False
                '--------------------------------------------------------------------------------------------------
            Else
                LimpioDatosCliente
            End If
            Set frmPersona = Nothing
        End If
    Else        'El cliente ingresado existe-------------------
        CargoDatosPersona
        BuscoComentariosAlerta RsAux!CliCodigo, labNombre.Caption, Alerta:=True
        RsAux.Close
        ValidoMayorDeEdad True, False
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
errBuscar:
    Screen.MousePointer = 0
    clsError.MuestroError "Ocurrió un error al cargar los datos del cliente."
End Sub

Private Sub BuscoGarantia(Cedula As String)
On Error GoTo errBuscar

    lGarantia.Caption = ""
    tGarantia.Tag = ""
    If Cedula = "" Then Exit Sub
    Screen.MousePointer = vbHourglass

    Cons = "Select CliCodigo, CliCIRuc, CliCategoria, CliDireccion,  CPersona.* From Cliente, CPersona " _
            & " Where CliCiRuc = '" & Cedula & "'" _
            & " And CliTipo = " & TipoCliente.Persona _
            & " And CliCodigo = CPeCliente"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)

    If RsAux.EOF Then
        Screen.MousePointer = 0
        RsAux.Close
        If MsgBox("No existe un cliente para la cédula de indentidad ingresada. Desea ingresarlo.", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
                    
            Set frmPersona = New MaCPersona
            frmPersona.pTipoLlamado = TipoLlamado.CreditoAClientes
            frmPersona.pClienteSeleccionado = 0
            frmPersona.pConyugeDe = 0
            frmPersona.tCi.Text = tGarantia.Text
            frmPersona.Show vbModal, Me
            
            If frmPersona.pClienteSeleccionado > 0 Then
                'Cargo los datos del Cliente recien ingresado---------------------------------------------
                Cons = "Select CliCodigo,CliCIRuc, CliCategoria, CliDireccion, CPersona.* From Cliente, CPersona " _
                    & " Where CliCodigo = " & frmPersona.pClienteSeleccionado _
                    & " And CliTipo = " & TipoCliente.Persona _
                    & " And CliCodigo = CPeCliente"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                CargoDatosGarantia
                BuscoComentariosAlerta RsAux!CliCodigo, lGarantia.Caption, Alerta:=True
                RsAux.Close
                ValidoMayorDeEdad False, True
                '--------------------------------------------------------------------------------------------------
            End If
            Set frmPersona = Nothing
        End If
    Else        'El cliente ingresado existe-------------------
        CargoDatosGarantia
        BuscoComentariosAlerta RsAux!CliCodigo, lGarantia.Caption, Alerta:=True
        RsAux.Close
        ValidoMayorDeEdad False, True
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
errBuscar:
    Screen.MousePointer = 0
    clsError.MuestroError "Ocurrió un error al cargar los datos de la garantía."
End Sub


Private Sub LimpioDatosCliente()

    labNombre.Caption = ""
    labDireccion.Caption = ""
    tGarantia.Text = FormatoCedula
    lGarantia.Caption = ""
    
    labNombre.Tag = vbNullString
    labDireccion.Tag = ""
    tGarantia.Tag = ""
    
    lTEdad.Caption = ""
    lGEdad.Caption = ""
    
End Sub


Private Sub cComentario_GotFocus()

    cComentario.SelStart = 0
    cComentario.SelLength = Len(cComentario.Text)
    
    Status.Panels(1).Text = "Ingrese un comentario para la solicitud."
    
End Sub

Private Sub cComentario_KeyPress(KeyAscii As Integer)

    cComentario.ListIndex = gIndice
    If KeyAscii = vbKeyReturn Then AccionGrabar
    
End Sub

Private Sub tEntrega_GotFocus()

    tEntrega.SelStart = 0
    tEntrega.SelLength = Len(tEntrega.Text)
    Status.Panels(1).Text = "Ingrese el valor de la entrega para el artículo seleccionado."
    
End Sub

Private Sub tEntrega_KeyPress(KeyAscii As Integer)

Dim aPlan As Long
Dim iAuxiliar As Currency

    If KeyAscii = vbKeyReturn And Trim(tEntrega.Text) <> "" Then
        If IsNumeric(tEntrega.Text) Then
            On Error GoTo errEntrega
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
            aPlan = Mid(lvVenta.SelectedItem.Key, InStr(lvVenta.SelectedItem.Key, "P") + 1, InStr(lvVenta.SelectedItem.Key, "A") - InStr(lvVenta.SelectedItem.Key, "P") - 1)

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
            lvVenta.SelectedItem.SubItems(6) = Redondeo(iAuxiliar / RsAux!TCuCantidad)
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
    Screen.MousePointer = 0
    clsError.MuestroError "Ocurrió un error al calcular los precios. Verifique los datos."
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

Private Sub tEntregaT_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tEntregaT.Text) <> "" Then
            If IsNumeric(tEntregaT.Text) And sDistribuir Then
                tEntregaT.Text = Format(tEntregaT.Text, "#,##0.00")
                DistribuirEntregas CCur(tEntregaT.Text)
            End If
            cPago.SetFocus
        End If
    End If
    
End Sub

Private Sub tGarantia_Change()

    tGarantia.Tag = ""
    lGarantia.Caption = ""
    lGEdad.Caption = ""
    
End Sub

Private Sub tGarantia_GotFocus()

    tGarantia.SelStart = 0
    tGarantia.SelLength = Len(tGarantia.Text)
    
    Status.Panels(1).Text = "Ingrese la cédula de identidad de la garantía."
    
End Sub

Private Sub tGarantia_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF4: If Shift = 0 Then BuscarGarantia                                   'Buscar Clientes
        Case vbKeyF3: IrAClienteNuevo False, True                 'Nuevo Cliente
        
        Case 93     'Boton Derecho
            If tGarantia.Tag = "0" Or tGarantia.Tag = "" Then MnuGFicha.Enabled = False Else: MnuGFicha.Enabled = True
            PopupMenu MnuMousseGarantia, , tGarantia.Left + (tGarantia.Width / 2), (tGarantia.Top + tGarantia.Height) - (tGarantia.Height / 2)
            
    End Select
    
End Sub

Private Sub tGarantia_KeyPress(KeyAscii As Integer)

On Error GoTo errBuscar

    If KeyAscii = vbKeyReturn Then
        Screen.MousePointer = 11
        Dim aCi As String
        'Valido la Cédula ingresada----------
        If Trim(tGarantia.Text) <> FormatoCedula Then
            aCi = QuitoFormatoCedula(tGarantia.Text)
            If Len(aCi) <> 8 Or Not CedulaValida(aCi) Then
                Screen.MousePointer = vbDefault
                lGarantia.Caption = ""
                MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
        End If
        'Busco el Cliente -----------------------
        If Trim(tGarantia.Text) <> FormatoCedula Then
            BuscoGarantia QuitoFormatoCedula(tGarantia.Text)
            If tGarantia.Tag <> vbNullString Then Foco cCuota
        Else
            Foco cCuota
        End If
        Screen.MousePointer = 0
    End If
    Exit Sub
    
errBuscar:
    Screen.MousePointer = 0
    clsError.MuestroError "Ocurrio un error al buscar la cédula de identidad."
End Sub

Private Sub tRuc_GotFocus()

    tRuc.SelStart = 0
    tRuc.SelLength = 15
    
    Status.Panels(1).Text = "Ingrese el número de R.U.C. de la empresa."
    
End Sub

Private Sub tRuc_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF4: If Shift = 0 Then BuscarClientes TipoCliente.Empresa        'Buscar Clientes
        Case vbKeyF3: IrAEmpresaNuevo                                  'Nuevo Cliente
        
        Case 93         'Boton Derecho
            If Mid(labNombre.Tag, 1, 1) = "E" Then MnuFichaEmpresa.Enabled = True Else: MnuFichaEmpresa.Enabled = False
            PopupMenu MnuMousseEmpresa, , tRuc.Left + (tRuc.Width / 2), (tRuc.Top + tRuc.Height) - (tRuc.Height / 2)
    End Select
    
End Sub

Private Sub tRuc_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        'Busco la Empresa------------------------
        If Trim(tRuc.Text) <> "" Then
            tCi.Text = FormatoCedula
            BuscoEmpresaRuc Trim(tRuc.Text)
            If labNombre.Tag <> vbNullString Then
                cCuota.SetFocus
            Else
                tCi.SetFocus
            End If
        Else
            If labNombre.Tag <> "" Then cCuota.SetFocus
        End If
    End If
    
End Sub

Private Sub BuscoEmpresaRuc(Codigo As String)

On Error GoTo errBuscar

    Screen.MousePointer = 11
    tCi.Text = FormatoCedula
    Cons = "Select CliCodigo,CliCIRuc, CliCategoria, CEmFantasia, CEmNombre, CliDireccion from Cliente, CEmpresa" _
            & " Where CliCiRuc = '" & Codigo & "'" _
            & " And CliTipo = " & TipoCliente.Empresa _
            & " And CliCodigo = CEmCliente"
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    LimpioDatosCliente
    If RsAux.EOF Then
        Screen.MousePointer = 0
        RsAux.Close
        If MsgBox("No existe una empresa para el número de R.U.C. ingresado. Desea ingresarla.", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
                
            Set frmEmpresa = New MaCEmpresa
            
            frmEmpresa.pClienteSeleccionado = 0
            frmEmpresa.pTipoLlamado = TipoLlamado.IngresoNuevo
            frmEmpresa.tRuc = tRuc.Text
            frmEmpresa.Show vbModal, Me
            
            If frmEmpresa.pClienteSeleccionado > 0 Then     'Si selecciono Cargo los datos de la Empresa
                Screen.MousePointer = 11
                Cons = "Select CliCodigo, CliCIRuc, CliCategoria, CEmFantasia, CEmNombre, CliDireccion from Cliente, CEmpresa" _
                    & " Where CliCodigo = " & frmEmpresa.pClienteSeleccionado _
                    & " And CliTipo = " & TipoCliente.Empresa _
                    & " And CliCodigo = CEmCliente"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                CargoDatosEmpresa
                BuscoComentariosAlerta RsAux!CliCodigo, labNombre.Caption, Alerta:=True
                RsAux.Close
            End If
            
            Set frmEmpresa = Nothing
        End If
    Else        'La empresa seleccionada Existe-----------
        CargoDatosEmpresa
        BuscoComentariosAlerta RsAux!CliCodigo, labNombre.Caption, Alerta:=True
        RsAux.Close
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errBuscar:
    Screen.MousePointer = 0
    clsError.MuestroError "Ocurrió un error al cargar los datos de la empresa."
End Sub

Private Sub CargoDatosPersona()

    LimpioTodaLaFicha False, True
    
    'ATENCION-------------------------------------------------------------------
    'En el tag del nombre me guardo el tipo de Cliente.
    'En el tag de dirección guardo la categoria de descuento del cliente.
    '---------------------------------------------------------------------------
    If Not IsNull(RsAux!CliCIRuc) Then tCi.Text = RetornoFormatoCedula(RsAux!CliCIRuc) Else: tCi.Text = FormatoCedula
    If Not IsNull(RsAux!CPERuc) Then tRuc.Text = Trim(RsAux!CPERuc) Else: tRuc.Text = ""
    
    labNombre.Caption = " " & ArmoNombre(Format(RsAux!CPeApellido1, "#"), Format(RsAux!CPeApellido2, "#"), Format(RsAux!CPeNombre1, "#"), Format(RsAux!CPeNombre2, "#"))
    labNombre.Tag = "P" & RsAux!CliCodigo
    
    If Not IsNull(RsAux!CliDireccion) Then labDireccion.Caption = " " & DireccionATextoSoloCalle(RsAux!CliDireccion)
    If Not IsNull(RsAux!CPERuc) Then tRuc.Text = Trim(RsAux!CPERuc)
    
    If Not IsNull(RsAux!CliCategoria) Then
        labDireccion.Tag = RsAux!CliCategoria
        CargoCuotas RsAux!CliCategoria
    Else
        CargoCuotas paCategoriaCliente
    End If
    
    If Not IsNull(RsAux!CPeFNacimiento) Then lTEdad.Caption = ((Date - RsAux!CPeFNacimiento) \ 365)

    
    
End Sub

Private Sub CargoDatosGarantia()
    
    tGarantia.Text = FormatoCedula
    tGarantia.Tag = ""
    lGarantia.Caption = ""
    lGEdad.Caption = ""
    
    If Not IsNull(RsAux!CliCIRuc) Then tGarantia.Text = RetornoFormatoCedula(RsAux!CliCIRuc)
    lGarantia.Caption = " " & ArmoNombre(Format(RsAux!CPeApellido1, "#"), Format(RsAux!CPeApellido2, "#"), Format(RsAux!CPeNombre1, "#"), Format(RsAux!CPeNombre2, "#"))
    tGarantia.Tag = RsAux!CliCodigo
    If Not IsNull(RsAux!CPeFNacimiento) Then lGEdad.Caption = ((Date - RsAux!CPeFNacimiento) \ 365)
    
End Sub
Private Sub CargoDatosEmpresa()

    LimpioTodaLaFicha False
    'ATENCION-------------------------------------------------------------------
    'En el tag del nombre me guardo el tipo de Cliente.
    'En el tag de dirección guardo la categoria de descuento del cliente.
    '---------------------------------------------------------------------------
    If Not IsNull(RsAux!CliCIRuc) Then tRuc.Text = Trim(RsAux!CliCIRuc)
    labNombre.Caption = " " & Trim(RsAux!CEmFantasia)
    If Not IsNull(RsAux!CEmNombre) Then labNombre.Caption = labNombre.Caption & " (" & Trim(RsAux!CEmNombre) & ")"
    
    labNombre.Tag = "E" & RsAux!CliCodigo
    If Not IsNull(RsAux!CliDireccion) Then labDireccion.Caption = " " & DireccionATextoSoloCalle(RsAux!CliDireccion)
    
    If Not IsNull(RsAux!CliCategoria) Then
        labDireccion.Tag = RsAux!CliCategoria
        CargoCuotas RsAux!CliCategoria
    Else
        CargoCuotas paCategoriaCliente
    End If
    

End Sub

Private Sub IrAEmpresa()
    
    On Error GoTo ErrIAE

    Screen.MousePointer = 11
    Set frmEmpresa = New MaCEmpresa
    
    If labNombre.Tag = vbNullString Then
        frmEmpresa.pClienteSeleccionado = 0
    Else
        If Left(labNombre.Tag, 1) = "E" Then
            frmEmpresa.pClienteSeleccionado = Mid(labNombre.Tag, 2, Len(labNombre.Tag) - 1)
        Else
            frmEmpresa.pClienteSeleccionado = 0
        End If
    End If
    
    frmEmpresa.pTipoLlamado = TipoLlamado.Visualizacion
    frmEmpresa.Show vbModal, Me
    Me.Refresh
    If frmEmpresa.pClienteSeleccionado > 0 Then
        LimpioDatosCliente
        tCi.Text = FormatoCedula
        
        Cons = "Select CliCodigo, CliCIRuc, CliCategoria, CEmFantasia, CEmNombre, CliDireccion from Cliente, CEmpresa" _
            & " Where CliCodigo = " & frmEmpresa.pClienteSeleccionado _
            & " And CliTipo = " & TipoCliente.Empresa _
            & " And CliCodigo = CEmCliente"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        CargoDatosEmpresa
        RsAux.Close
    
    Else    'Si lo Borro
         If frmPersona.pClienteSeleccionado = 0 Then LimpioDatosCliente
    End If
    Set frmEmpresa = Nothing
    Screen.MousePointer = 0
    Exit Sub

ErrIAE:
    Screen.MousePointer = 0
    clsError.MuestroError "Ocurrió un error al acceder a los datos de la empresa."
End Sub

Private Sub IrAEmpresaNuevo()

    On Error GoTo errNuevo
    Screen.MousePointer = 11
    LimpioDatosCliente
    Set frmEmpresa = New MaCEmpresa
            
    frmEmpresa.pClienteSeleccionado = 0
    frmEmpresa.pTipoLlamado = TipoLlamado.IngresoNuevo
    frmEmpresa.tRuc = tRuc.Text
    frmEmpresa.Show vbModal, Me
    Me.Refresh
    
    If frmEmpresa.pClienteSeleccionado > 0 Then     'Si selecciono Cargo los datos de la Empresa
        Screen.MousePointer = 11
        Cons = "Select CliCodigo, CliCIRuc, CliCategoria, CEmFantasia, CEmNombre, CliDireccion from Cliente, CEmpresa" _
            & " Where CliCodigo = " & frmEmpresa.pClienteSeleccionado _
            & " And CliTipo = " & TipoCliente.Empresa _
            & " And CliCodigo = CEmCliente"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        CargoDatosEmpresa
        BuscoComentariosAlerta RsAux!CliCodigo, labNombre.Caption, Alerta:=True
        RsAux.Close
    End If
    
    Set frmEmpresa = Nothing
    Screen.MousePointer = 0
    Exit Sub
    
errNuevo:
    Screen.MousePointer = 0
    clsError.MuestroError "Ocurrió un error al ingresar la nueva empresa."
End Sub

Private Sub IrACliente(Titular As Boolean, Garantia As Boolean, Optional AIngresarFNac As Boolean = False)

    On Error GoTo ErrIAC
    Screen.MousePointer = 11
    Set frmPersona = New MaCPersona
    If AIngresarFNac Then
        frmPersona.pTipoLlamado = TipoLlamado.ClienteParaIngresoFNacimiento
    Else
        frmPersona.pTipoLlamado = TipoLlamado.Visualizacion
    End If
    
    If Titular Then         'Llamado desde CI TITULAR----------------------------------------------------------
        If labNombre.Tag = vbNullString Then
            frmPersona.pClienteSeleccionado = 0
        Else
            If Left(labNombre.Tag, 1) = "P" Then
                frmPersona.pClienteSeleccionado = Mid(labNombre.Tag, 2, Len(labNombre.Tag) - 1)
            Else
                frmPersona.pClienteSeleccionado = 0     'El que esta es una empresa.----------------
            End If
        End If
    End If '-------------------------------------------------------------------------------------------------------------
    
    If Garantia Then        'Llamado desde CI GARANTIA----------------------------------------------------------
        If tGarantia.Tag <> "" Then
            frmPersona.pClienteSeleccionado = tGarantia.Tag
        Else
            frmPersona.pClienteSeleccionado = 0
        End If
    End If '--------------------------------------------------------------------------------------------------------------
    
    frmPersona.pConyugeDe = 0
    frmPersona.Show vbModal, Me
    Me.Refresh
    
    If frmPersona.pClienteSeleccionado > 0 Then
        Cons = "Select CliCodigo, CliCIRuc, CliCategoria, CliDireccion, CPersona.* From Cliente, CPersona " _
            & " Where CliCodigo = " & frmPersona.pClienteSeleccionado _
            & " And CliTipo = " & TipoCliente.Persona _
            & " And CliCodigo = CPeCliente"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If Titular Then CargoDatosPersona    'Cargo datos del Cliente Titular
        
        If Garantia Then CargoDatosGarantia    'Cargo Datos del Cliente Garantia
       
        RsAux.Close
    Else
        'Si los Borró
        If frmPersona.pClienteSeleccionado = 0 Then
            If Titular Then LimpioTodaLaFicha False, True: tCi.Text = FormatoCedula
            
            If Garantia Then
                tGarantia.Text = FormatoCedula
                tGarantia.Tag = ""
                lGarantia.Caption = ""
                lGEdad.Caption = ""
            End If
        End If
    End If
    Set frmPersona = Nothing
    
    If (Titular And labNombre.Caption <> "") Or (Garantia And lGarantia.Caption <> "") Then
        ValidoMayorDeEdad Titular, Garantia
        If Titular And labNombre.Caption <> "" Then
            If tGarantia.Tag = "" Or tGarantia.Tag = "0" Then tGarantia.SetFocus Else Foco cCuota
        End If
        If Garantia And lGarantia.Caption <> "" Then cCuota.SetFocus
    End If
    
    Screen.MousePointer = 0
    Exit Sub

ErrIAC:
    Screen.MousePointer = 0
    clsError.MuestroError "Ocurrió un error al acceder a los datos del cliente."
End Sub

Private Sub IrAClienteNuevo(Titular As Boolean, Garantia As Boolean)

    On Error GoTo ErrIAC
    Screen.MousePointer = 11
    Set frmPersona = New MaCPersona
    frmPersona.pTipoLlamado = TipoLlamado.CreditoAClientes
    frmPersona.pClienteSeleccionado = 0
    frmPersona.pConyugeDe = 0
    frmPersona.Show vbModal, Me
    Me.Refresh
    
    If frmPersona.pClienteSeleccionado > 0 Then
        Cons = "Select CliCodigo, CliCIRuc, CliCategoria, CliDireccion, CPersona.* From Cliente, CPersona " _
            & " Where CliCodigo = " & frmPersona.pClienteSeleccionado _
            & " And CliTipo = " & TipoCliente.Persona _
            & " And CliCodigo = CPeCliente"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If Titular Then CargoDatosPersona     'Cargo datos del Cliente Titular
        
        If Garantia Then CargoDatosGarantia   'Cargo Datos del Cliente Garantia
        RsAux.Close
        
    End If
    
    Set frmPersona = Nothing
    
    If (Titular And labNombre.Caption <> "") Or (Garantia And lGarantia.Caption <> "") Then
        ValidoMayorDeEdad Titular, Garantia
        If Titular And labNombre.Caption <> "" Then     '
            If tGarantia.Tag = "" Or tGarantia.Tag = "0" Then tGarantia.SetFocus Else Foco cCuota
        End If
        If Garantia And lGarantia.Caption <> "" Then cCuota.SetFocus
    End If
    Screen.MousePointer = 0
    Exit Sub

ErrIAC:
    Screen.MousePointer = 0
    clsError.MuestroError "Ocurrió un error al acceder a los datos del cliente."
End Sub

Private Sub BuscoArticuloXNombre()
On Error GoTo ErrBAN

    Screen.MousePointer = 11

    Cons = "Select ArtId, ArtNombre from Articulo" _
           & " Where ArtNombre LIKE '" & cArticulo.Text & "%'" _
           & " Order by ArtNombre"
    
    sAyuda = True
    LiAyuda.pSeleccionado = 0
    LiAyuda.pConsulta = Cons
    LiAyuda.pEncabezado = "Artículo"
    LiAyuda.pEncabezadoValores = "3000"
    LiAyuda.pFormato = "#"
    LiAyuda.pAlineacion = "0"
    
    LiAyuda.lAyuda.HideColumnHeaders = True
    LiAyuda.lAyuda.Width = 3700
    LiAyuda.Width = LiAyuda.lAyuda.Width + 80
    
    LiAyuda.Show vbModal
    Me.Refresh
    
    sAyuda = False
    cArticulo.Clear
    If LiAyuda.pSeleccionado > 0 Then
        Dim aArticulo As Long
        Cons = "Select * From Articulo Where ArtId = " & LiAyuda.pSeleccionado
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        aArticulo = RsAux!ArtCodigo
        RsAux.Close
        BuscoArticuloxCodigo aArticulo
    Else
        tUnitario.Text = ""
    End If
    Screen.MousePointer = 0
    Exit Sub
    
ErrBAN:
    Screen.MousePointer = 0
    clsError.MuestroError "Ocurrió un error al buscar el artículo."
End Sub

Private Sub BuscoArticuloxCodigo(Codigo As Long)

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
    Cons = "Select * From Articulo Where ArtCodigo = " & Codigo
    Set RsArt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsArt.EOF Then
        If IsNull(RsArt!ArtHabilitado) Or UCase(RsArt!ArtHabilitado) <> "S" Then
            MsgBox "El artículo no está habilitado para la venta. Consulte.", vbExclamation, Trim(RsArt!ArtNombre)
            RsArt.Close
            Screen.MousePointer = 0
            Exit Sub
        End If
    Else
        MsgBox "El artículo ingresado no existe. Verifique los datos.", vbExclamation, "ATENCIÓN"
        RsArt.Close
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    'Verifico si el articulo está en la lista---------------------------------------------------------------------------------
    If Ingresado(RsArt!ArtID, cCuota.ItemData(cCuota.ListIndex)) Then
        Screen.MousePointer = 0
        MsgBox "El artículo seleccionado ya fue ingresado. Para modificarlo edítelo.", vbExclamation, "ATENCIÓN"
        RsArt.Close
        Exit Sub
    End If
    '--------------------------------------------------------------------------------------------------------------------------
    cArticulo.AddItem Trim(RsArt!ArtNombre)
    cArticulo.ItemData(cArticulo.NewIndex) = RsArt!ArtID
    cArticulo.ListIndex = 0
    Codigo = RsArt!ArtID
    RsArt.Close
    '------------------------------------------------------------------------------------------------------------------
    
    If Not sConEntrega Then         'PROCESO PLAN SIN ENTREGA------------------------------------------------
        'Saco el valor de la cuota financiado
        Cons = "Select PViPrecio, PViHabilitado, PViPlan From PrecioVigente" _
                & " Where PVIArticulo = " & Codigo _
                & " And PViMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                & " And PViTipoCuota = " & cCuota.ItemData(cCuota.ListIndex)
        Set RsArt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        
        If Not RsArt.EOF Then       'Hay Precios Grabados
            If Not (RsArt!PViHabilitado) Then
                Screen.MousePointer = 0
                MsgBox "El artículo no está habilitado a la venta para la moneda seleccionada. Consulte.", vbCritical, "ATENCIÓN"
                RsArt.Close
                cArticulo.Clear
                Exit Sub
            End If
            lSubTotalF.Tag = Format(RsArt!PViPrecio, "#,##0")                          'Precio de Unitario Financiaciado
            tUnitario.Text = Format(RsArt!PViPrecio / cCuota.Tag, "#,##0.00")     'Valor Cuota Finanaciado
            tCantidad.Tag = RsArt!PViPlan
        
        Else            'No Hay Precios Grabados
            'Si Hay precio Contado NO LO VENDO
            Cons = "Select PViPrecio, PViHabilitado, PViPlan From PrecioVigente" _
                & " Where PVIArticulo = " & Codigo _
                & " And PViMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                & " And PViTipoCuota = " & paTipoCuotaContado _
                & " And PViHabilitado = 1"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux.EOF Then   'Hay contado...no se vende
                Screen.MousePointer = 0
                MsgBox "El artículo no se puede vender con el plan seleccionado." & Chr(vbKeyReturn) & "Hay precios contado, pero financiados no.", vbCritical, "ATENCIÓN"
                cArticulo.Clear
            Else
                'Como no hay ni Contado, ni Credito pido el valor de la cuota
                tCantidad.Tag = paPlanPorDefecto    'El plan no lo necesito
            End If
            RsAux.Close
        End If
        RsArt.Close
    
    Else                                            'PROCESO PLAN CON ENTREGA------------------------------------------------
        
        'Busco el precio contado del articulo para el plan seleccionado-------------------------
        Cons = "Select PViPrecio, PViHabilitado, PViPlan From PrecioVigente" _
                & " Where PVIArticulo = " & Codigo _
                & " And PViMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                & " And PViTipoCuota = " & paTipoCuotaContado _
                & " And PViHabilitado = 1"
        Set RsArt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        
        'Si no hay CONTADO, hay que pedirlo para el plan que ingreso ---> voy a trabajar con el plan por defecto
        If Not RsArt.EOF Then
            tUnitario.Tag = RsArt!PViPrecio    'Precio Unitario Contado
            tUnitario.Text = Format(RsArt!PViPrecio, FormatoMonedaP)   'Precio Unitario Contado
            tCantidad.Tag = RsArt!PViPlan
        Else
            tCantidad.Tag = paPlanPorDefecto
        End If
        RsArt.Close
        
        'Valido que exista un coeficiente para el calculo (TipoCuota, Plan, Moneda) Ya sea para el plan ingresado o el por defecto
        Cons = "Select * from Coeficiente" _
                & " Where CoePlan = " & tCantidad.Tag _
                & " And CoeTipoCuota = " & cCuota.ItemData(cCuota.ListIndex) _
                & " And CoeMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If RsAux.EOF Then    'Si NO hay coeficientes NO SE VENDE
            Screen.MousePointer = 0
            MsgBox "No existe un coeficiente para el cálculo de cuotas. Consulte.", vbExclamation, "ATENCIÓN"
            cArticulo.Clear
            tUnitario.Text = "": tUnitario.Tag = ""
        End If
        RsAux.Close
    End If
    Screen.MousePointer = 0
    Exit Sub
    
ErrBAC:
    Screen.MousePointer = 0
    clsError.MuestroError "Ocurrió un error al buscar el artículo."
End Sub

Private Sub BuscoArticuloxCodigo2(Codigo As Long)

'   lSubTotalF.Tag  = Precio Unitario Financiado
'   cCuota.Tag       = Cantidad de Cuotas
'   tCantidad.Tag   = Plan
'   tUnitario.tag = Precio Unitario Contado        ------ y el en Text va el valor de la cuota

On Error GoTo ErrBAC

    Screen.MousePointer = 11
    'Saco el Precio contado del Articulo--------------------------------------------------------------------------
    Cons = "Select ArtID, ArtNombre, PViPrecio, PViHabilitado, PViPlan From Articulo, PrecioVigente" _
            & " Where ArtCodigo = " & Codigo _
            & " And ArtID = PViArticulo " _
            & " And ArtHabilitado = 'S'" _
            & " And PViMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
            & " And PViTipoCuota = " & paTipoCuotaContado
    Set RsArt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    cArticulo.Clear
    tUnitario.Text = ""
    tUnitario.Tag = ""
    lSubTotalF.Tag = ""
    If RsArt.EOF Then
        RsArt.Close
        'Busco los precios por el plan definido por defecto en paPlanPorDefecto
        Screen.MousePointer = 0
        MsgBox "No se encontró el artículo o no tiene precios para el plan seleccionado. Consulte", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    tUnitario.Tag = RsArt!PViPrecio    'Precio Unitario Contado
    If sConEntrega Then tUnitario.Text = Format(RsArt!PViPrecio, FormatoMonedaP)   'Precio Unitario Contado
    tCantidad.Tag = RsArt!PViPlan
    '------------------------------------------------------------------------------------------------------------------
    
    'Verifico si el articulo está en la lista---------------------------------------------------------------------------------
    If Ingresado(RsArt!ArtID, cCuota.ItemData(cCuota.ListIndex)) Then
        Screen.MousePointer = 0
        MsgBox "El artículo seleccionado ya fue ingresado, para modificarlo edítelo.", vbExclamation, "ATENCIÓN"
        tUnitario.Text = ""
        tUnitario.Tag = ""
        RsArt.Close
        Exit Sub
    End If
    '--------------------------------------------------------------------------------------------------------------------------
    
    cArticulo.AddItem Trim(RsArt!ArtNombre)
    cArticulo.ItemData(cArticulo.NewIndex) = RsArt!ArtID
    cArticulo.ListIndex = 0
    RsArt.Close
    
    If Not sConEntrega Then
        'Saco el valor de la cuota financiado
        Cons = "Select ArtID, ArtNombre, PViPrecio, PViHabilitado, PViPlan From Articulo, PrecioVigente" _
                & " Where ArtCodigo = " & Codigo _
                & " And ArtID = PViArticulo " _
                & " And ArtHabilitado = 'S'" _
                & " And PViMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                & " And PViTipoCuota = " & cCuota.ItemData(cCuota.ListIndex)
        Set RsArt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If RsArt.EOF Then
            tUnitario.Text = ""
            Screen.MousePointer = 0
            MsgBox "No hay precios ingresados para el plan seleccionado. Consulte", vbExclamation, "ATENCIÓN"
            RsArt.Close
            cArticulo.Clear
            Exit Sub
        End If
        If Not (RsArt!PViHabilitado) Then
            Screen.MousePointer = 0
            tUnitario.Text = ""
            MsgBox "El artículo no está habilitado para la venta con la moneda seleccionada, consulte.", vbCritical, "ATENCIÓN"
            RsArt.Close
            cArticulo.Clear
            Exit Sub
        End If
        lSubTotalF.Tag = Format(RsArt!PViPrecio, "#,##0")                          'Precio de Unitario Financiaciado
        tUnitario.Text = Format(RsArt!PViPrecio / cCuota.Tag, "#,##0.00")     'Valor Cuota Finanaciado
        RsArt.Close
    Else
        'Valido que exista un coeficiente para el calculo (TipoCuota, Plan, Moneda)
        Cons = "Select * from Coeficiente" _
                & " Where CoePlan = " & tCantidad.Tag _
                & " And CoeTipoCuota = " & cCuota.ItemData(cCuota.ListIndex) _
                & " And CoeMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If RsAux.EOF Then
            Screen.MousePointer = 0
            MsgBox "No hay ingresado un coeficiente para el cálculo de cuotas. Consulte.", vbExclamation, "ATENCIÓN"
            tUnitario.Text = ""
            cArticulo.Clear
        End If
        RsAux.Close
    End If
    Screen.MousePointer = 0
    Exit Sub
    
ErrBAC:
    Screen.MousePointer = 0
    clsError.MuestroError "Ocurrió un error al buscar el artículo."
End Sub

Private Sub tUnitario_Change()
    lSubTotalF.Caption = ""
End Sub

Private Sub tUnitario_GotFocus()

Dim aUnitarioF As Currency

    Status.Panels(1).Text = "Precio contado unitario del artículo."
    tUnitario.SelStart = 0
    tUnitario.SelLength = Len(tUnitario.Text)
    
    If sConEntrega Or Trim(tCantidad.Text) = "" Or Not IsNumeric(tCantidad.Text) _
    Or cArticulo.ListIndex = -1 Then Exit Sub
    
    If lSubTotalF.Tag <> vbNullString Then      'En el TAG Tengo el Precio financiado del articulo
        aUnitarioF = BuscoDescuentoCliente(cArticulo.ItemData(cArticulo.ListIndex), Val(labDireccion.Tag), CCur(lSubTotalF.Tag), _
                            Val(tCantidad.Text), cArticulo.Text, cCuota.ItemData(cCuota.ListIndex))
        tUnitario.Text = Redondeo(aUnitarioF / CCur(cCuota.Tag))
        lSubTotalF.Caption = Format(((tUnitario.Text * CCur(cCuota.Tag)) * CCur(tCantidad.Text)), FormatoMonedaP)
    End If

End Sub

Private Sub tUnitario_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And IsNumeric(tUnitario.Text) And cMoneda.ListIndex > -1 Then
        If IsNumeric(tUnitario.Text) Then
            If Not sConEntrega Then
                'Hay que calcular los totales -- por si cambió la cuota.
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
                                                          Articulo As String, Plan As Long) As String

Dim RsBDC As rdoResultset
Dim aRetorno As String

    On Error GoTo errDescuento
    aRetorno = Redondeo(Unitario)
    
    If CodCatCliente > 0 Then
    
        Cons = "Select CDTPorcentaje, AFaCantidadD From ArticuloFacturacion, CategoriaDescuento" _
                & " Where AFaArticulo = " & CodArticulo _
                & " And AFaCategoriaD = CDtCatArticulo " _
                & " And CDtCatCliente = " & CodCatCliente _
                & " And CDtCatPlazo = " & Plan
            
        Set RsBDC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        
        If Not RsBDC.EOF Then
            If Not IsNull(RsBDC!AFaCantidadD) Then
                If RsBDC!AFaCantidadD <= Cantidad Then
                    aRetorno = Redondeo(Unitario - (Unitario * RsBDC(0)) / 100)
                Else
                    If MsgBox("La cantidad no llega a la mímima (" & RsBDC!AFaCantidadD & ") para aplicar el descuento. " & Chr(vbKeyReturn) _
                                & "Desea aplicar el descuento correspondiente.", vbQuestion + vbYesNo, Trim(Articulo)) = vbYes Then
                        aRetorno = Redondeo(Unitario - (Unitario * RsBDC(0)) / 100)
                    End If
                End If
            End If
        End If
        
        RsBDC.Close
    End If
    
    BuscoDescuentoCliente = aRetorno
    Exit Function

errDescuento:
    Screen.MousePointer = 0
    clsError.MuestroError "Ocurrió un error al procesar los descuentos."
    BuscoDescuentoCliente = aRetorno
End Function

Private Sub InsertoFila()

    On Error GoTo ErrIF
    'Valido los campos para insertar la linea de articulo-----------------------------------------------------
    If cCuota.ListIndex = -1 Then
        MsgBox "Debe seleccionar el tipo de financiación.", vbExclamation, "ATENCIÓN"
        Foco cCuota
        Exit Sub
    End If
    If cArticulo.ListIndex = -1 Then
        MsgBox "Debe seleccionar un artículo.", vbExclamation, "ATENCIÓN"
        Foco cArticulo
        Exit Sub
    End If
    If Not IsNumeric(tCantidad.Text) Then
        MsgBox "La cantidad ingresada no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tCantidad
        Exit Sub
    End If
    If Not Val(tCantidad.Text) > 0 Then
        MsgBox "La cantidad ingresada no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tCantidad
        Exit Sub
    End If
    If Not IsNumeric(tUnitario.Text) Then
        MsgBox "El precio unitario ingresado no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tUnitario
        Exit Sub
    Else
        If CCur(tUnitario.Text) < 0 Then
            MsgBox "No se puede facturar artículos con costo negativo.", vbExclamation, "ATENCIÓN"
            Foco tUnitario
            Exit Sub
        End If
    End If
    '-----------------------------------------------------------------------------------------------------------------------
    
    If sConEntrega Then
        Set itmX = lvVenta.ListItems.Add(, "E" & cCuota.ItemData(cCuota.ListIndex) & "P" & tCantidad.Tag & "A" & cArticulo.ItemData(cArticulo.ListIndex), cCuota.Text)
        itmX.SubItems(8) = tUnitario.Tag     'Unitario Contado (para controlar cambio de precios con entrega)
        itmX.SubItems(3) = Format(tUnitario.Text, "#,##0.00")                     'Contado
    Else
        Set itmX = lvVenta.ListItems.Add(, "N" & cCuota.ItemData(cCuota.ListIndex) & "P" & tCantidad.Tag & "A" & cArticulo.ItemData(cArticulo.ListIndex), cCuota.Text)
        itmX.SubItems(3) = Format(tUnitario.Tag, "#,##0.00")                     'Contado
    End If
    
    itmX.SubItems(1) = Trim(tCantidad.Text)
    itmX.SubItems(2) = Trim(cArticulo.Text)

    itmX.SubItems(4) = IVAArticulo(cArticulo.ItemData(cArticulo.ListIndex))
    
    If Trim(lSubTotalF.Caption) <> "" Then
        'Cuota
        itmX.SubItems(6) = Format(CCur(tUnitario.Text) * CCur(tCantidad.Text), FormatoMonedaP)
        
        'Ajusto el subtotal con lo que me da la cuota (SubTotal)
        itmX.SubItems(7) = lSubTotalF.Caption   'Total Financiado
        
        itmX.SubItems(8) = lSubTotalF.Tag     'Unitario Financiado
        
        TotalesSumo CCur(itmX.SubItems(7)), CCur(itmX.SubItems(4))
    End If
    
    LimpioRenglon
    cArticulo.Clear
    cCuota.SetFocus
    If lvVenta.ListItems.Count > 0 Then
        cMoneda.Enabled = False
        MnuEmitir.Enabled = True
    Else
        cMoneda.Enabled = True
        MnuEmitir.Enabled = False
    End If
    HabilitoEntrega
    Exit Sub
    
ErrIF:
    clsError.MuestroError "Ocurrió un error al insertar el renglón."
End Sub

Private Function IVAArticulo(lngCodigo As Long)

    Cons = "Select IVAPorcentaje From ArticuloFacturacion, TipoIva " _
        & " Where AFaArticulo = " & lngCodigo _
        & " And AFaIVA = IVACodigo"
        
    Set RsArt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If RsArt.EOF Then
        IVAArticulo = 0
    Else
        IVAArticulo = Format(RsArt(0), "#0.00")
    End If
    RsArt.Close

End Function


Private Function Ingresado(Articulo As Long, Cuota As Long)

    If lvVenta.ListItems.Count > 0 Then
        Ingresado = False
        For Each itmX In lvVenta.ListItems
            If Mid(itmX.Key, 2, InStr(itmX.Key, "P") - 2) = Cuota And Mid(itmX.Key, InStr(itmX.Key, "A") + 1, Len(itmX.Key)) = Articulo Then
                Ingresado = True
                Exit Function
            End If
        Next
    Else
        Ingresado = False
    End If

End Function

Private Sub TotalesResto(Total As Currency, Iva As Currency)

    labIva.Caption = Format(CCur(labIva.Caption) - (Total - (Total / (1 + Iva / 100))), "#,##0.00")
    labTotal.Caption = Format(CCur(labTotal.Caption) - Total, "#,##0.00")
    labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIva.Caption), "#,##0.00")
    
End Sub

Private Sub TotalesSumo(Total As Currency, Iva As Currency)

    labIva.Caption = Format(CCur(labIva.Caption) + Total - (Total / (1 + Iva / 100)), "#,##0.00")
    labTotal.Caption = Format(CCur(labTotal.Caption) + Total, "#,##0.00")
    labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIva.Caption), "#,##0.00")
        
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
    If ControloDatos Then
        If MsgBox("Confirma almacenar los datos ingresados en la solicitud.", vbQuestion + vbYesNo, "GRABAR") = vbYes Then
            If labDireccion.Tag = "" Then labDireccion.Tag = "0"
            
            Dim aDiferenciaPrecios As Currency
            aDiferenciaPrecios = ControloPrecios(CLng(labDireccion.Tag))
            If aDiferenciaPrecios <> 0 Then
                'Llamo al registro del Suceso-------------------------------------------------------------
                InSuceso.pNombreSuceso = "Cambio de Precios"
                InSuceso.Show vbModal, Me
                If InSuceso.pUsuario = 0 Then Exit Sub  'Abortó el ingreso del suceso
                Me.Refresh
            End If
            
            GrabarSolicitud aDiferenciaPrecios
            
        End If
    End If
    Exit Sub

errGrabar:
    Screen.MousePointer = 0
    clsError.MuestroError "Ocurrió un error al procesar los datos."
End Sub

Private Function ControloPrecios(CategoriaCliente As Long) As Currency

Dim aItmx As ListItem
Dim RsCon As rdoResultset
Dim aUnitario As Currency
Dim aDiferencia As Currency     'Diferencia de Precios

    On Error GoTo errControl
    Screen.MousePointer = 11
    ControloPrecios = 0: aDiferencia = 0
    
    For Each aItmx In lvVenta.ListItems
        
        If Trim(aItmx.SubItems(8)) <> "" Then       'Si tenia precio para controlar
            If Trim(aItmx.SubItems(5)) = "" Then    '(5)= Entrega -- Si el Plan es con entrega no hago el control
                'SOLO PARA PLANES SIN ENTREGA!!!!!!
                aUnitario = CCur(aItmx.SubItems(8))
                If aUnitario * CCur(aItmx.SubItems(1)) <> CCur(aItmx.SubItems(7)) Then
                    '2 posibilidades --> o hay descuento o cambió el precio
                    'Veo Si hay descuentos
                    If CategoriaCliente > 0 Then        'Hago las consultas para comparar precios
                        
                        Cons = "Select CDTPorcentaje, AFaCantidadD, TCuCantidad From ArticuloFacturacion, CategoriaDescuento, TipoCuota" _
                                & " Where AFaArticulo = " & ArticuloDeLaClave(aItmx.Key) _
                                & " And AFaCategoriaD = CDtCatArticulo " _
                                & " And CDtCatCliente = " & CategoriaCliente _
                                & " And CDtCatPlazo = " & CuotaDeLaClave(aItmx.Key) & " And CDtCatPlazo = TCuCodigo "
                            
                        Set RsCon = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                        If Not RsCon.EOF Then
                            'Unitario - (Unitario * %Dto) / 100
                            aUnitario = Redondeo(CCur(aItmx.SubItems(8)) - (CCur(aItmx.SubItems(8)) * RsCon(0)) / 100)
                            'Para que no de diferencia, hay que Unitario = Redondeo(Unitario/CCuotas)
                            aUnitario = CCur(Redondeo(aUnitario / RsCon!TCuCantidad)) * RsCon!TCuCantidad
                        End If
                        RsCon.Close
                    End If
                    'OJO en el Sbt(7) está el financiado sin descuentos
                    'Si el unitario * cantidad <> Subtotal ------> HayQueGrabarSuceso
                    If aUnitario * CCur(aItmx.SubItems(1)) <> CCur(aItmx.SubItems(7)) Then
                        aDiferencia = aDiferencia + (CCur(aItmx.SubItems(7)) - aUnitario * CCur(aItmx.SubItems(1)))
                    End If
                End If
            Else        'Plan Con entrega Comparo si modificó el contado
                aUnitario = CCur(aItmx.SubItems(8))     'Initario
                If aUnitario <> CCur(aItmx.SubItems(3)) Then
                    aDiferencia = aDiferencia + (CCur(aItmx.SubItems(3)) - aUnitario)
                End If
            End If
        End If
    Next
    
    ControloPrecios = aDiferencia
    
    Screen.MousePointer = 0
    Exit Function

errControl:
    Screen.MousePointer = 11
    clsError.MuestroError "Ocurrió un error al controlar los precios de artículos."
End Function

Private Function ControloDatos() As Boolean
Dim Suma As Currency

    ControloDatos = False
    
    If labNombre.Tag = vbNullString Then
        MsgBox "No se puede emitir una solicitud sin seleccionar un cliente.", vbExclamation, "ATENCIÓN"
        Foco tCi
        Exit Function
    End If
    If cMoneda.ListIndex = -1 Then
        MsgBox "Se debe seleccionar una moneda para emitir la solicitud.", vbExclamation, "ATENCIÓN"
        cMoneda.Enabled = True
        Foco cMoneda
        Exit Function
    End If
    If lvVenta.ListItems.Count = 0 Then
        MsgBox "Debe ingresar los artículos para la solicitud.", vbExclamation, "ATENCIÓN"
        Foco cArticulo
        Exit Function
    End If
    
    If tGarantia.Tag <> vbNullString And tGarantia.Text = FormatoCedula Then
        MsgBox "La garantía seleccionada debe presentar la cédula de identidad.", vbExclamation, "ATENCIÓN"
        Foco tGarantia
        Exit Function
    End If
    
    'Valido la suma de los Subtotales contra el Total -----------------------------
    Suma = 0
    For Each itmX In lvVenta.ListItems
        If Trim(itmX.SubItems(7)) = "" Then
            MsgBox "Se debe ingresar el valor de la entrega para las financiaciones con entrega.", vbExclamation, "ATENCIÓN"
            Foco tEntregaT
            Exit Function
        End If
        Suma = Suma + CCur(itmX.SubItems(7))
    Next
    If Suma <> CCur(labTotal.Caption) Then
        MsgBox "La suma total no coincide con la suma de la lista, verifique.", vbCritical, "ATENCIÓN"
        cArticulo.SetFocus
        Exit Function
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
                Foco tEntregaT
                Exit Function
            End If
        Else
            MsgBox "La suma de las entregas no coincide con el valor ingresado, verifique.", vbExclamation, "ATENCIÓN"
            Foco tEntregaT
            Exit Function
        End If
    End If
    '-----------------------------------------------------------------------------------
    
    If cPago.ListIndex = -1 Then
        MsgBox "Se debe seleccionar la forma de pago de la solicitud.", vbExclamation, "ATENCIÓN"
        Foco cPago
        Exit Function
    End If
    
    If Trim(tVendedor.Tag) = "" Or tVendedor.Tag = "0" Or tVendedor.Tag = 0 Then
        MsgBox "Debe ingresar el dígito del vendedor.", vbExclamation, "ATENCIÓN"
        Foco tVendedor
        Exit Function
    End If
    
    If tUsuario.Tag = vbNullString Then
        MsgBox "Debe ingresar el dígito de usuario.", vbExclamation, "ATENCIÓN"
        Foco tUsuario
        Exit Function
    End If
    
    'Si el pago es con Cheque Dif. controlo que las cuotas no superen el parametro = paCantidadMaxCheques
    If cPago.ItemData(cPago.ListIndex) = TipoPagoSolicitud.ChequeDiferido Then
        For Each itmX In lvVenta.ListItems
            If Trim(itmX.SubItems(5)) = "" Then
                'Subtotal / Valor Cuota = Cant. Cuotas
                If CCur(itmX.SubItems(7)) / CCur(itmX.SubItems(6)) > paCantidadMaxCheques Then
                    MsgBox "Hay cuotas que superan la cantidad máxima de cheques diferidos." & Chr(vbKeyReturn) & "Cambie la forma de pago.", vbExclamation, "ATENCIÓN"
                    Foco cPago
                    Exit Function
                End If
            Else
                '(Subtotal - Entrega) / Valor Cuota = Cant. Cuotas
                If (CCur(itmX.SubItems(7)) - CCur(itmX.SubItems(5))) / CCur(itmX.SubItems(6)) > paCantidadMaxCheques Then
                    MsgBox "Hay cuotas que superan la cantidad máxima de cheques diferidos." & Chr(vbKeyReturn) & "Cambie la forma de pago.", vbExclamation, "ATENCIÓN"
                    Foco cPago
                    Exit Function
                End If
            End If
        Next
    End If
    
    ControloDatos = True
    
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
    clsError.MuestroError "Ocurrio un error inesperado."
    BuscoUsuario = 0
    
End Function

Private Sub GrabarSolicitud(Optional DiferenciaPrecios As Currency = 0)

Dim aSolicitud As Long

    Screen.MousePointer = vbHourglass
    On Error GoTo ErrGFR
    FechaDelServidor
    cBase.BeginTrans
    
    On Error GoTo ErrResumo
    
    'Cargo los datos de la SOLICITUD-------------------------------------------------------------------------------------
    Cons = "Select * From Solicitud Where SolCodigo = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.AddNew
    
    RsAux!SolCliente = Mid(labNombre.Tag, 2, Len(labNombre.Tag))
    RsAux!SolFecha = Format(gFechaServidor, FormatoFH)
    RsAux!SolTipo = TipoSolicitud.AlMostrador
    RsAux!SolProceso = TipoResolucionSolicitud.Manual
    RsAux!SolEstado = EstadoSolicitud.Pendiente
    
    If Trim(tGarantia.Tag) <> "" Then RsAux!SolGarantia = tGarantia.Tag
    RsAux!SolFormaPago = cPago.ItemData(cPago.ListIndex)
    RsAux!SolMoneda = cMoneda.ItemData(cMoneda.ListIndex)
    If Trim(cComentario.Text) <> "" Then RsAux!SolComentarioS = Trim(cComentario.Text)
    RsAux!SolUsuarioS = tUsuario.Tag
    RsAux!SolSucursal = paCodigoDeSucursal
    RsAux!SolVendedor = CLng(tVendedor.Tag)
    
    RsAux.Update
    RsAux.Close
    '---------------------------------------------------------------------------------------------------------------------------
    
    'Saco el numero de la solicitud--------------------------------------------------
    Cons = "SELECT MAX(SolCodigo) From Solicitud"
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
        RsAux!RSoTipoCuota = CuotaDeLaClave(itmX.Key)   ' Mid(itmX.Key, 2, InStr(itmX.Key, "P") - 2)
        RsAux!RSoArticulo = ArticuloDeLaClave(itmX.Key)     'Mid(itmX.Key, InStr(itmX.Key, "A") + 1, Len(itmX.Key))
        If Trim(itmX.SubItems(5)) <> "" Then RsAux!RSoValorEntrega = CCur(itmX.SubItems(5))
        RsAux!RSoValorCuota = CCur(itmX.SubItems(6))
        RsAux!RSoCantidad = itmX.SubItems(1)
        
        RsAux.Update
    Next
    RsAux.Close
    '-------------------------------------------------------------------------------------------------
    
    If DiferenciaPrecios <> 0 Then
        aTexto = "Solicitud de Crédito Nº " & aSolicitud
        RegistroSuceso gFechaServidor, TipoSuceso.ModificacionDePrecios, paCodigoDeTerminal, InSuceso.pUsuario, 0, Descripcion:=aTexto, Defensa:=Trim(InSuceso.pDefensa), Valor:=DiferenciaPrecios
    End If
    
    cBase.CommitTrans                               'Fin TRANSACCION----------------------------------------------!!!!!!!!!!!!!!!!!!!
    
    Screen.MousePointer = vbDefault
    tCi.Text = FormatoCedula
    LimpioTodaLaFicha True
    Exit Sub
    
ErrGFR:
    Screen.MousePointer = vbDefault
    clsError.MuestroError "Ocurrió un error al iniciar la transacción."
    Exit Sub

ErrResumo:
    Resume ErrRelajo
    
ErrRelajo:
    Screen.MousePointer = vbDefault
    cBase.RollbackTrans
    clsError.MuestroError "Ocurrió un error al realizar la solicitud."
End Sub

Private Function ValidoBuscarArticulo()

    ValidoBuscarArticulo = False
    
    If cMoneda.ListIndex = -1 Then
        MsgBox "Se debe seleccionar una moneda para realizar la solicitud.", vbCritical, "ATENCIÓN"
        Foco cMoneda
        Exit Function
    End If

    If cCuota.ListIndex = -1 Then
        MsgBox "Se debe seleccionar el tipo de financiación para realizar la solicitud.", vbCritical, "ATENCIÓN"
        Foco cCuota
        Exit Function
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
            sHay = True
            Exit For
        End If
    Next
    If Not sHay Then
        Screen.MousePointer = 0
        MsgBox "No hay financiaciones con entrega para realizar la distribución.", vbInformation, "ATENCIÓN"
        tEntregaT.Text = ""
        Exit Sub
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
        Foco tEntregaT
        Exit Sub
    End If
    '--------------------------------------------------------------------------------------------------
    
    Dim itmP As ListItem
    For Each itmX In lvVenta.ListItems
        If Left(itmX.Key, 1) = "E" Then
            'Veo Si ya hice la distribucion
            If Trim(itmX.SubItems(5)) = "" Then
                'Con el Total distriubuyo el porcentaje de la entrega
                For Each itmP In lvVenta.ListItems
                    If itmX.Text = itmP.Text Then   'El mismo Tipo de Cuota
                        itmP.SubItems(5) = Format(((CCur(itmP.SubItems(3)) * CCur(itmP.SubItems(1)) * 100) / aTotal) * ValorAEntregar / 100, "#,##0.00")
                        
                        'El valor de la cuota es el (Precio Contado - Entrega) * Coeficiente ----- Coeficiente (Plan, TCuota, Moneda)
                        Cons = "Select * from Coeficiente, TipoCuota" _
                            & " Where CoePlan = " & Mid(itmP.Key, InStr(itmP.Key, "P") + 1, InStr(itmP.Key, "A") - InStr(itmP.Key, "P") - 1) _
                            & " And CoeTipoCuota = " & Mid(itmP.Key, 2, InStr(itmP.Key, "P") - 2) _
                            & " And CoeMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                            & " And CoeTipocuota = TCuCodigo"
                        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                        'Calculo lo que queda por pagar ((P.contado * Cantidad)- Entega * Coeficiente)
                        iAuxiliar = ((CCur(itmP.SubItems(3)) * CCur(itmP.SubItems(1))) - CCur(itmP.SubItems(5))) * RsAux!CoeCoeficiente
                        
                        'Veo si tiene descuento
                        iAuxiliar = CCur(BuscoDescuentoCliente(ArticuloDeLaClave(itmP.Key), Val(labDireccion.Tag), iAuxiliar, CCur(itmP.SubItems(1)), itmP.SubItems(2), CuotaDeLaClave(itmP.Key)))
                        
                        'Valor de Cada Cuota
                        itmP.SubItems(6) = Redondeo(iAuxiliar / RsAux!TCuCantidad)
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
    Screen.MousePointer = 0
    clsError.MuestroError "Ocurrió un error al realizar la distribución de la entrega."
End Sub

Private Sub HabilitoEntrega()

Dim sHay As Boolean

    sHay = False
    For Each itmX In lvVenta.ListItems
        If Left(itmX.Key, 1) = "E" Then
            sHay = True
            Exit For
        End If
    Next
    
    If sHay Then
        tEntregaT.Enabled = True
        tEntregaT.BackColor = Obligatorio
    Else
        tEntregaT.Enabled = False
        tEntregaT.BackColor = Inactivo
    End If
    
End Sub

Private Sub IngresoDeEntrega(Valor As Boolean)

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
            If MsgBox("El cliente seleccionado no tiene ingresada la fecha de nacimiento. Para realizar la solicitud debe ingresarla, desea hacerlo.", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
                LimpioDatosCliente
            Else
                IrACliente True, False, True
            End If
        Else
            If CLng(lTEdad.Caption) < paMayorDeEdad Then
                Screen.MousePointer = 0
                MsgBox "El cliente seleccionado es menor de edad. No puede solicitar créditos.", vbExclamation, "ATENCIÓN"
                LimpioDatosCliente
            End If
        End If
    Else
        If Trim(lGEdad.Caption) = "" Then
            Screen.MousePointer = 0
            If MsgBox("El cliente seleccionado no tiene ingresada la fecha de nacimiento. Para actuar como garante debe ingresarla, desea hacerlo.", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
                tGarantia.Text = FormatoCedula
                tGarantia.Tag = ""
                lGarantia.Caption = ""
                lGEdad.Caption = ""
            Else
                IrACliente False, True, True
            End If
        Else
            If CLng(lGEdad.Caption) < paMayorDeEdad Then
                Screen.MousePointer = 0
                MsgBox "El cliente seleccionado es menor de edad. No puede presentarse como garantía del crédito.", vbExclamation, "ATENCIÓN"
                tGarantia.Text = FormatoCedula
                tGarantia.Tag = ""
                lGarantia.Caption = ""
                lGEdad.Caption = ""
            End If
        End If
    End If
    
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

Private Function ArticuloDeLaClave(Clave As String) As Long
    ArticuloDeLaClave = CLng(Mid(Clave, InStr(Clave, "A") + 1, Len(Clave)))
End Function

Private Function CuotaDeLaClave(Clave As String) As Long
    CuotaDeLaClave = CLng(Mid(Clave, 2, InStr(Clave, "P") - 2))
End Function

