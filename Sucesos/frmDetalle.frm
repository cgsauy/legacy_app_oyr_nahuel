VERSION 5.00
Begin VB.Form frmDetalle 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalle del Suceso"
   ClientHeight    =   4185
   ClientLeft      =   2790
   ClientTop       =   3630
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton bAnterior 
      Caption         =   "<< &Anterior"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   20
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton bSiguiente 
      Caption         =   "&Siguiente >>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1500
      TabIndex        =   19
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Información General"
      ForeColor       =   &H00000080&
      Height          =   2055
      Left            =   60
      TabIndex        =   6
      Top             =   360
      Width           =   6675
      Begin VB.TextBox tDefensa 
         Appearance      =   0  'Flat
         Height          =   555
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   1440
         Width           =   6495
      End
      Begin VB.Label lUsrAutoriza 
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4560
         TabIndex        =   27
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lAutoriza 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   3360
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Terminal:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Motivos/Defensa:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   2235
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha/Hora:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lFecha 
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4560
         TabIndex        =   10
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lTerminal 
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lDescripcion 
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   720
         Width           =   5295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Relacionados"
      ForeColor       =   &H00000080&
      Height          =   1215
      Left            =   60
      TabIndex        =   1
      Top             =   2520
      Width           =   6675
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lCliente 
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1080
         TabIndex        =   24
         Top             =   960
         Width           =   5415
      End
      Begin VB.Label lValor 
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1080
         TabIndex        =   18
         Top             =   720
         Width           =   5415
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Documento:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Artículo:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lDocumento 
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label lArticulo 
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   480
         Width           =   5415
      End
   End
   Begin VB.CommandButton bSalir 
      Caption         =   "Sali&r"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5760
      TabIndex        =   0
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lId 
      Caption         =   "S/D"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   900
      TabIndex        =   22
      Top             =   75
      Width           =   915
   End
   Begin VB.Label Label6 
      Caption         =   "CODIGO:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   90
      Width           =   855
   End
   Begin VB.Label lTipo 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "S/D"
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
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   60
      Width           =   4575
   End
End
Attribute VB_Name = "frmDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public prm_Suceso As Long
 
Dim aTexto As String
Dim aQRows As Long
Dim aRow As Long

Dim aDocumento As Long, aArticulo As Long, aCliente As Long

Private Sub bAnterior_Click()

Dim aCodigo As Long

    On Error GoTo errSuceso
       
    If aRow > 1 Then
        Screen.MousePointer = 11
        LimpioFicha
        aRow = aRow - 1
        frmSuceso.vsConsulta.Select aRow, 0
        frmSuceso.vsConsulta.TopRow = aRow
        
        prm_Suceso = frmSuceso.vsConsulta.Cell(flexcpValue, aRow, 0)
        CargoDatosSuceso prm_Suceso
        
        Screen.MousePointer = 0
        bSiguiente.Enabled = True
        
        If aRow = 1 Then bAnterior.Enabled = False
        
    Else
        bAnterior.Enabled = False
    End If
    
    Exit Sub

errSuceso:
    Screen.MousePointer = 0
End Sub

Private Sub bSalir_Click()
    Unload Me
End Sub

Private Sub bSiguiente_Click()

    On Error GoTo errSuceso
    If aRow < aQRows Then
    
        LimpioFicha
        Screen.MousePointer = 11
        aRow = aRow + 1
        frmSuceso.vsConsulta.Select aRow, 0
        frmSuceso.vsConsulta.TopRow = aRow
        
        prm_Suceso = frmSuceso.vsConsulta.Cell(flexcpValue, aRow, 0)
        CargoDatosSuceso prm_Suceso
        
        Screen.MousePointer = 0
        bAnterior.Enabled = True
        
        If aRow = aQRows Then bSiguiente.Enabled = False
    Else
        bSiguiente.Enabled = False
    End If                                      '------------------------------------------------------------

    Exit Sub

errSuceso:
    Screen.MousePointer = 0
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub


Private Sub Form_Load()
    On Error Resume Next
    
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    aQRows = frmSuceso.vsConsulta.Rows - 1
    aRow = frmSuceso.vsConsulta.Row
    
    If prm_Suceso <> 0 Then CargoDatosSuceso prm_Suceso
    
    If aRow = 1 Then bAnterior.Enabled = False
    If aRow = aQRows Then bSiguiente.Enabled = False
    
End Sub


Private Sub CargoDatosSuceso(Suceso As Long)


    aDocumento = 0: aArticulo = 0: aCliente = 0
    lId.Caption = Format(Suceso, "#,##0")
    
    cons = "Select * from Suceso, Terminal, Usuario, TipoSuceso" _
            & " Where SucCodigo = " & Suceso _
            & " And SucTipo = TSuCodigoSistema" _
            & " And SucTerminal *= TerCodigo" _
            & " And SucUsuario = UsuCodigo"
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not rsAux.EOF Then
        lTipo.Caption = Trim(rsAux!TSuNombre)
        lFecha.Caption = Format(rsAux!SucFecha, "Ddd d-Mmm yyyy hh:mm:ss")
        If Not IsNull(rsAux!SucDescripcion) Then lDescripcion.Caption = Trim(rsAux!SucDescripcion)
        If Not IsNull(rsAux!TerNombre) Then lTerminal.Caption = Trim(rsAux!TerNombre)
        lUsuario.Caption = Trim(rsAux!UsuApellido1) & ", " & Trim(rsAux!UsuNombre1)
        If Not IsNull(rsAux!SucValor) Then lValor.Caption = Format(rsAux!SucValor, FormatoMonedaP)
        
        If Not IsNull(rsAux!SucDefensa) Then tDefensa.Text = Trim(rsAux!SucDefensa)
        
        If Not IsNull(rsAux!SucDocumento) Then aDocumento = rsAux!SucDocumento
        If Not IsNull(rsAux!SucArticulo) Then aArticulo = rsAux!SucArticulo
        If Not IsNull(rsAux!SucCliente) Then aCliente = rsAux!SucCliente
        
        'SucAutoriza SucVerificado
        If Not IsNull(rsAux!SucAutoriza) Then
            lUsrAutoriza.Caption = DatosUsuario(rsAux!SucAutoriza, "UsuIdentificacion")
            If Not IsNull(rsAux!SucVerificado) Then
                If rsAux!SucVerificado Then lAutoriza.Caption = "Autorizado:" Else lAutoriza.Caption = "Para Autorizar:"
            End If
        End If
    End If
    
    rsAux.Close
    
    If aDocumento <> 0 Then CargoDatosDocumento aDocumento Else lDocumento.Caption = "S/D"
    If aArticulo <> 0 Then CargoDatosArticulo aArticulo Else lArticulo.Caption = "S/D"
    If aCliente <> 0 Then CargoDatosCliente aCliente Else lCliente.Caption = "S/D"
    
End Sub

Private Function DatosUsuario(idUsr As Long, idCampo As String) As Variant
On Error GoTo errUser

    DatosUsuario = ""
    Dim arrData() As String, arrCampo() As String
    arrData() = Split(miConexion.UserInfo(idUsr, 0), vbCrLf)
    For I = LBound(arrData) To UBound(arrData)
        arrCampo = Split(arrData(I), "=")
        If LCase(arrCampo(0)) = LCase(idCampo) Then
            DatosUsuario = arrCampo(1)
            Exit For
        End If
    Next
    
    

errUser:
End Function

Private Sub CargoDatosDocumento(Codigo As Long)
    On Error GoTo errCDoc
    
    aTexto = ""
    cons = "Select * from Documento, Sucursal" _
            & " Where DocCodigo =  " & Codigo _
            & " And DocSucursal = SucCodigo"
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then
    
        aTexto = Trim(rsAux!SucAbreviacion) & " "
        
        Select Case rsAux!DocTipo
            Case TipoDocumento.Contado: aTexto = aTexto & "Contado "
            Case TipoDocumento.Credito: aTexto = aTexto & "Crédito "
            Case TipoDocumento.ReciboDePago: aTexto = aTexto & "Recibo Pago "
            Case TipoDocumento.NotaCredito: aTexto = aTexto & "N/Crédito "
            Case TipoDocumento.NotaDevolucion: aTexto = aTexto & "N/Contado "
            Case TipoDocumento.NotaEspecial: aTexto = aTexto & "N/Especial "
            
            Case Else: aTexto = ""
        End Select
        
        If aTexto <> "" Then
            aTexto = aTexto & Trim(rsAux!DocSerie) & " " & rsAux!DocNumero & " " _
                                    & "(" & Format(rsAux!DocFecha, "d/mm/yy hh:mm") & ")"
                                    
            If aCliente = 0 Then aCliente = rsAux!DocCliente
        End If
    End If
    rsAux.Close
    
    If aTexto = "" Then aTexto = "id_Referencia " & Codigo
    
    lDocumento.Caption = aTexto

errCDoc:
End Sub


Private Sub CargoDatosArticulo(Codigo As Long)
    On Error GoTo errCArt
    
    cons = "Select * from Articulo Where ArtId = " & Codigo
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then
        lArticulo.Caption = Format(rsAux!ArtCodigo, "(#,000,000)") & " " & Trim(rsAux!ArtNombre)
    End If
    rsAux.Close

errCArt:
End Sub


Private Sub CargoDatosCliente(Codigo As Long)
    On Error GoTo errCCliente
    aTexto = ""
    cons = "Select CliCiRuc, CliTipo, Nombre =  (RTrim(CPeNombre1) + RTrim(' ' + CPeNombre2)+' ' + RTrim(CPeApellido1)) + RTrim(' ' + CPeApellido2)" _
           & " From Cliente, CPersona Where CPeCliente = " & Codigo & " And CliCodigo = CPeCliente" _
                        & " Union All " _
           & " Select CliCiRuc, CliTipo, Nombre = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')')" _
           & " From Cliente, CEmpresa Where CEmCliente = " & Codigo & " And CliCodigo = CEmCliente"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux!CliCiRuc) Then
            If rsAux!CliTipo = 1 Then aTexto = clsGeneral.RetornoFormatoCedula(rsAux!CliCiRuc)
            If rsAux!CliTipo = 2 Then aTexto = clsGeneral.RetornoFormatoRuc(rsAux!CliCiRuc)
            aTexto = "(" & Trim(aTexto) & ") "
        End If
        
        aTexto = aTexto & Trim(rsAux!Nombre)
    End If
    rsAux.Close
    
    If Trim(aTexto) <> "" Then lCliente.Caption = aTexto

errCCliente:
End Sub

Private Sub LimpioFicha()
   
    lAutoriza.Caption = "Autoriza:": lUsrAutoriza.Caption = "S/D"
    lFecha.Caption = "S/D"
    lDescripcion.Caption = "S/D"
    lTerminal.Caption = "S/D"
    lUsuario.Caption = "S/D"
    lValor.Caption = "S/D"
    
    tDefensa.Text = ""
    
    lDocumento.Caption = "S/D"
    lArticulo.Caption = "S/D"
    lCliente.Caption = "S/D"
    
End Sub

