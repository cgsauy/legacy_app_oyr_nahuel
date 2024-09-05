VERSION 5.00
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEntProgramada 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrega Programada"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEntProgramada.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5880
      TabIndex        =   15
      Top             =   3360
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   4440
      TabIndex        =   14
      Top             =   3360
      Width           =   1275
   End
   Begin AACombo99.AACombo cboLocal 
      Height          =   315
      Left            =   1320
      TabIndex        =   7
      Top             =   2760
      Width           =   3255
      _ExtentX        =   5741
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
   Begin VB.TextBox txtNumero 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Text            =   "B 88888888"
      Top             =   1080
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker txtFecha 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
      Format          =   56033283
      CurrentDate     =   40417
   End
   Begin VB.PictureBox picTitulo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   7470
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   7470
      Begin VB.Image Image1 
         Height          =   480
         Left            =   6600
         Picture         =   "frmEntProgramada.frx":06EA
         Top             =   195
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Entregas programadas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000040C0&
      X1              =   7200
      X2              =   240
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label hlkDocumento 
      BackStyle       =   0  'Transparent
      Caption         =   "Contado B 546088"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2520
      MouseIcon       =   "frmEntProgramada.frx":0BB1
      MousePointer    =   99  'Custom
      TabIndex        =   12
      ToolTipText     =   "Acceder a detalle de factura"
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Coordinar entrega"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Emisión"
      Height          =   255
      Left            =   5280
      TabIndex        =   10
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblDocEmision 
      BackStyle       =   0  'Transparent
      Caption         =   "15/12/2020 14:55"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   6000
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblDocCliente 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      DataField       =   "Juan Pêr"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   1440
      Width           =   5775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Local"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
End
Attribute VB_Name = "frmEntProgramada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private prmLocales As String
Private Function CargoLocales() As Boolean
'Controlo aquellos que son vitales si no los cargue finalizo la app.
On Error GoTo errCP
    
    'Parametros a cero--------------------------
    Dim sLocales As String
    
    Cons = "Select RTRIM(ParTexto) Texto from Parametro Where ParNombre IN('EntProgramada_Locales')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        sLocales = RsAux("Texto")
    End If
    RsAux.Close
    
    If sLocales = "" Then Exit Function
    
    Cons = "SELECT SucCodigo, SucAbreviacion FROM Sucursal WHERE SucCodigo IN(" & sLocales & ")"
    CargoCombo Cons, cboLocal
    
    If cboLocal.ListCount > 0 Then cboLocal.ListIndex = 0
    Exit Function
    
errCP:
     clsGeneral.OcurrioError "Error al cargar los locales.", Err.Description
     CargoLocales = False
End Function
Sub LimpioFichaDocumento()
    lblDocCliente.Caption = ""
    lblDocEmision.Caption = ""
    txtNumero.Tag = ""
    hlkDocumento.Caption = ""
    cboLocal.Text = ""
    txtFecha.Value = Now
    EstadoControles False
End Sub

Sub EstadoControles(ByVal habilitados As Boolean)
    cmdOK.Enabled = habilitados
    txtFecha.Enabled = habilitados
    cboLocal.Enabled = habilitados
End Sub


Private Sub cboLocal_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Sub AccionGrabar()
    
    If CDate(txtFecha.Value) < Date Then
        MsgBox "La fecha debe ser igual o mayor a hoy.", vbExclamation, "ATENCIÓN"
        txtFecha.SetFocus
        Exit Sub
    End If
    
    If cboLocal.ListIndex = -1 Then
        MsgBox "Es necesario indicar el local donde se retira la mercadería.", vbExclamation, "ATENCIÓN"
        On Error Resume Next
        cboLocal.SetFocus
        Exit Sub
    End If
    
    
    On Error GoTo errSave
    Cons = "SELECT * FROM EntregaProgramada " & _
        "WHERE EPrDocumento = " & txtNumero.Tag & " AND EPrTipo = 2 AND EPrLocal = " & cboLocal.ItemData(cboLocal.ListIndex)
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        
        If MsgBox("¿Confirma MODIFICAR la entrega programada para el " & Format(RsAux("EPrFecha"), "dd/MM/yyyy hh:nn:ss") & "?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
            RsAux.Edit
        Else
            RsAux.Close
            Exit Sub
        End If
    Else
        If MsgBox("¿Confirma programar la entrega de la mercadería" & vbCrLf & " para el " & txtFecha.Value & vbCrLf & " en el local " & cboLocal.Text & "?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
            RsAux.AddNew
        Else
            RsAux.Close
            Exit Sub
        End If
    End If
    
    RsAux("EPrFecha") = Format(txtFecha.Value, "yyyy/mm/dd hh:nn:ss")
    RsAux("EPrTipo") = 2
    RsAux("EPrDocumento") = txtNumero.Tag
    RsAux("EPrLocal") = cboLocal.ItemData(cboLocal.ListIndex)
    RsAux.Update
    
    RsAux.Close
    
    MsgBox "Datos almacenados", vbInformation, "ATENCIÓN"
    LimpioFichaDocumento

    Exit Sub
errSave:
    clsGeneral.OcurrioError "Error al almacenar la entrega.", Err.Description, "Error al grabar"
    Screen.MousePointer = 0
End Sub

Private Sub cmdOK_Click()
    AccionGrabar
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CargoLocales
    txtNumero.Text = ""
    LimpioFichaDocumento
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    cBase.Close
End Sub

Private Sub hlkDocumento_Click()
On Error Resume Next
    If Val(txtNumero.Tag) > 0 Then Shell App.Path & "\detalle de factura.exe " & txtNumero.Tag, vbNormalFocus
End Sub

Private Sub txtFecha_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyReturn Then cboLocal.SetFocus
End Sub


Private Sub txtNumero_Change()
    If Val(txtNumero.Tag) <> 0 Then LimpioFichaDocumento
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Val(txtNumero.Tag) <> 0 Then Exit Sub
        If Trim(txtNumero.Text) = "" Then Exit Sub
        
        LimpioFichaDocumento
   
        Dim mDSerie As String, mDNumero As Long
        Dim adQ As Integer, adCodigo As Long, adTexto As String
        
        adTexto = Trim(txtNumero.Text)
        If InStr(adTexto, "-") <> 0 Then
            mDSerie = Mid(adTexto, 1, InStr(adTexto, "-") - 1)
            mDNumero = Val(Mid(adTexto, InStr(adTexto, "-") + 1))
        Else
            adTexto = Replace(adTexto, " ", "")
            If IsNumeric(Mid(adTexto, 2, 1)) Then
                mDSerie = Mid(adTexto, 1, 1)
                mDNumero = Val(Mid(adTexto, 2))
            Else
                mDSerie = Mid(adTexto, 1, 2)
                mDNumero = Val(Mid(adTexto, 3))
            End If
        End If
        txtNumero.Text = UCase(mDSerie) & "-" & mDNumero
        
        Screen.MousePointer = 11
        adQ = 0: adTexto = ""
        Dim fecha As Date
        Dim cliente As String
        Dim doc As String
        
        Cons = "Select DocCodigo, DocFecha as Fecha" & _
                   ", Case CliTipo When 1 Then RTrim(RTrim(CPeApellido1) + ' ' + RTrim(IsNull(CPeApellido2, ''))) + ', ' + RTrim(RTrim(CPeNombre1) + ' ' + IsNull(CPeNombre2, '')) " & _
                                "When 2 Then RTrim(CEmFantasia) + IsNull('(' + RTrim(CEmNombre) + ')' , '') END Cliente " & _
                   ", CASE DocTipo WHEN 1 Then 'Contado' WHEN 2 Then 'Crédito' WHEN 3 Then 'Nota Dev' WHEN 4 then 'Nota Créd' WHEN 5 then 'Recibo' When 6 Then 'Remito' When 10 Then 'Nota Esp' ELSE '' END + " & _
                   " ' ' + rTrim(DocSerie) + '-' + rtrim(Convert(Varchar(6), DocNumero)) as Número" & _
                   " From Documento INNER JOIN Cliente ON DocCliente = CliCodigo " & _
                                "LEFT OUTER JOIN CPersona ON CPeCliente = CliCodigo " & _
                                "LEFT OUTER JOIN CEmpresa ON CEmCliente = CliCodigo " & _
                   " Where DocSerie = '" & mDSerie & "'" & _
                   " And DocNumero = " & mDNumero & _
                   " And DocTipo IN (" & TipoDocumento.Contado & ", " & TipoDocumento.Credito & ") " & _
                   " And DocAnulado = 0"
                                                   
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            adCodigo = RsAux!DocCodigo
            cliente = RsAux("Cliente")
            fecha = RsAux("Fecha")
            doc = RsAux("Número")
            adQ = 1
            RsAux.MoveNext: If Not RsAux.EOF Then adQ = 2
        End If
        RsAux.Close
        
        Select Case adQ
            Case 2
                Dim miLDocs As New clsListadeAyuda
                adCodigo = miLDocs.ActivarAyuda(cBase, Cons, 8100, 1)
                Me.Refresh
                If adCodigo > 0 Then
                    adCodigo = miLDocs.RetornoDatoSeleccionado(0)
                    fecha = miLDocs.RetornoDatoSeleccionado(1)
                    cliente = miLDocs.RetornoDatoSeleccionado(2)
                    doc = miLDocs.RetornoDatoSeleccionado(3)
                End If
                Set miLDocs = Nothing
                
        End Select
        
        If adCodigo > 0 Then
            
            txtNumero.Tag = adCodigo: hlkDocumento.Caption = doc
            lblDocEmision.Caption = Format(fecha, "dd/mm/yyyy")
            lblDocCliente.Caption = cliente
            
            EstadoControles True
            cboLocal.ListIndex = 0
            txtFecha.SetFocus
            
        Else
            hlkDocumento.Caption = " No Existe !!"
        End If
        Screen.MousePointer = 0
    End If
    
    Exit Sub
errDoc:
    clsGeneral.OcurrioError "Error al buscar el documento.", Err.Description
    Screen.MousePointer = 0
End Sub
