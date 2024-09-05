VERSION 5.00
Begin VB.Form frmAutorizaSuceso 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorizar suceso"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAutorizaSuceso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton butCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5280
      TabIndex        =   18
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton butNoAutorizar 
      Caption         =   "No autorizar"
      Height          =   375
      Left            =   3720
      TabIndex        =   17
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton butAutorizar 
      Caption         =   "Autorizar"
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label lblFecha 
      BackStyle       =   0  'Transparent
      Caption         =   "dd/MM/yyyy"
      Height          =   255
      Left            =   4800
      TabIndex        =   20
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   3840
      TabIndex        =   19
      Top             =   2760
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6735
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label lblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "lblCliente"
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   4800
      Width           =   5775
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002C511E&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   4440
      Width           =   3375
   End
   Begin VB.Label lblImporte 
      BackStyle       =   0  'Transparent
      Caption         =   "8,888,888.00"
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Importe:"
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblDocumento 
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Documento:"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lblCliente 
      BackStyle       =   0  'Transparent
      Caption         =   "lblCliente"
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   3600
      Width           =   5775
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002C511E&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle:"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblDefensa 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
      Height          =   615
      Left            =   1800
      TabIndex        =   6
      Top             =   2040
      Width           =   4575
   End
   Begin VB.Label lblDescripcion 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1560
      Width           =   4575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblTipo 
      BackStyle       =   0  'Transparent
      Caption         =   "Aportes a cuenta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo:"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Suceso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002C511E&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Autorizar sucesos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmAutorizaSuceso.frx":08CA
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmAutorizaSuceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public IDSuceso As Long

Private Sub Form_Load()
On Error GoTo errLFrm
    Screen.MousePointer = 11
    'Cargo la info del suceso.
    Dim rsS As rdoResultset
    Dim sQy As String
    
    lblCliente.Caption = ""
    lblDefensa.Caption = ""
    lblDescripcion.Caption = ""
    lblDocumento.Caption = ""
    lblImporte.Caption = ""
    lblFecha.Caption = ""
    lblTipo.Caption = ""
    lblUsuario.Caption = ""
    
    sQy = "SELECT CliCodigo, " & _
        "CASE CliTipo WHEN 1 THEN RTrim(RTrim(CPeApellido1) + ' ' + RTrim(IsNull(CPeApellido2, ''))) + ', ' + RTrim(RTrim(CPeNombre1) + ' ' + IsNull(CPeNombre2, '')) " & _
        "WHEN 2 Then RTrim(CEmFantasia) + IsNull('(' + RTrim(CEmNombre) + ')' , '') END Cliente, " & _
        "UsuIdentificacion, TerNombre, DocCodigo, DocSerie + '-' + Convert(varchar(6), DocNumero) Documento, TDoNombre, " & _
        "SucValor, SucFecha, TSuNombre " & _
        "FROM Suceso " & _
        "INNER JOIN cgsa.dbo.Cliente ON SucCliente = CliCodigo " & _
        "LEFT OUTER JOIN cgsa.dbo.CPersona ON CPeCliente = CliCodigo " & _
        "LEFT OUTER JOIN cgsa.dbo.CEmpresa ON CEmCliente = CliCodigo " & _
        "LEFT OUTER JOIN Documento ON DocCodigo = SucDocumento " & _
        "LEFT OUTER JOIN TipoDocumento ON TDoDocumento = DocTipo " & _
        "INNER JOIN Usuario ON SucUsuario = UsuCodigo " & _
        "LEFT OUTER JOIN Terminal ON SucTerminal = TerCodigo " & _
        "INNER JOIN TipoSuceso ON SucTipo = TSuCodigo " & _
        "WHERE SucCodigo = " & Me.IDSuceso

    Set rsS = cBase.OpenResultset(sQy, rdOpenForwardOnly, rdConcurReadOnly)
    If Not rsS.EOF Then
        If Not IsNull(rsS("Cliente")) Then lblCliente.Caption = rsS("Cliente")
        lblDefensa.Caption = Trim(rsS("SucDefensa"))
        lblDescripcion.Caption = Trim(rsS("SucDescripcion"))
        If Not IsNull(rsS("DocCodigo")) Then
            lblDocumento.Caption = RTrim(rsS("TDoNombre")) & " " & rsS("Documento")
        End If
        If Not IsNull(rsS("SucValor")) Then lblImporte.Caption = Format(rsS("SucValor"), "#,##0.00")
        lblFecha.Caption = Format(rsS("SucFecha"), "dd/MM/yy HH:nn")
        lblTipo.Caption = Trim(rsS("TSuNombre"))
        lblUsuario.Caption = rsS("UsuIdentificacion") & " Terminal: " & Trim(rsS("TerNombre"))
    End If
    rsS.Close
    
    Screen.MousePointer = 0
    Exit Sub
errLFrm:
    clsGeneral.OcurrioError "Error al inicializar el formulario.", Err.Description, "Autorizar suceso"
    Screen.MousePointer = 0
End Sub

