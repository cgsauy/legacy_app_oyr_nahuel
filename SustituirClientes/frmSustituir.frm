VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSustituir 
   Caption         =   "Sustituir Clientes"
   ClientHeight    =   6840
   ClientLeft      =   2205
   ClientTop       =   2055
   ClientWidth     =   7605
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSustituir.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   7605
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   6585
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12885
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Left            =   120
      ScaleHeight     =   4035
      ScaleWidth      =   6315
      TabIndex        =   11
      Top             =   420
      Width           =   6375
      Begin VB.Frame Frame1 
         Caption         =   "Cliente a Borrar"
         ForeColor       =   &H00800000&
         Height          =   1575
         Left            =   60
         TabIndex        =   14
         Top             =   60
         Width           =   6195
         Begin VB.TextBox tBCiRuc 
            Height          =   285
            Left            =   4320
            TabIndex        =   5
            Top             =   260
            Width           =   1755
         End
         Begin VB.TextBox tBId 
            Height          =   285
            Left            =   960
            TabIndex        =   7
            Top             =   260
            Width           =   975
         End
         Begin VB.TextBox tBTexto 
            BackColor       =   &H000000FF&
            ForeColor       =   &H00FFFFFF&
            Height          =   915
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   600
            Width           =   5955
         End
         Begin VB.Label Label1 
            Caption         =   "Id Cli&ente:"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   300
            Width           =   795
         End
         Begin VB.Label Label2 
            Caption         =   "C&I o RUC:"
            Height          =   255
            Left            =   3480
            TabIndex        =   4
            Top             =   300
            Width           =   795
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cliente Sustituto"
         ForeColor       =   &H00800000&
         Height          =   1575
         Left            =   60
         TabIndex        =   12
         Top             =   2040
         Width           =   6195
         Begin VB.TextBox tSCiRuc 
            Height          =   285
            Left            =   4320
            TabIndex        =   1
            Top             =   240
            Width           =   1755
         End
         Begin VB.TextBox tSId 
            Height          =   285
            Left            =   960
            TabIndex        =   3
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox tSTexto 
            BackColor       =   &H00008000&
            ForeColor       =   &H00FFFFFF&
            Height          =   915
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   585
            Width           =   5955
         End
         Begin VB.Label Label3 
            Caption         =   "Id &Cliente:"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   285
            Width           =   795
         End
         Begin VB.Label Label4 
            Caption         =   "C&I o RUC:"
            Height          =   255
            Left            =   3480
            TabIndex        =   0
            Top             =   285
            Width           =   795
         End
      End
      Begin VB.CommandButton bBorrar 
         Caption         =   "Borrar"
         Height          =   315
         Left            =   5340
         TabIndex        =   8
         Top             =   3660
         Width           =   915
      End
      Begin VB.Label lEstado 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
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
         Left            =   2100
         TabIndex        =   16
         Top             =   1740
         Width           =   4035
      End
   End
   Begin MSComctlLib.TabStrip tab1 
      Height          =   1815
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3201
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Datos a Sustituir"
            Key             =   "datos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " Evolución  "
            Key             =   "proceso"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.TextBox tReporte 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   3915
      Left            =   2820
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   4155
   End
End
Attribute VB_Name = "frmSustituir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type typCliente
    Codigo As Long
    Tipo As Integer
    CiRuc As String
    Nombre1 As String
    Nombre2 As String
    Apellido1 As String
    Apellido2 As String
    FNac As String
    idDireccion As Long
    Direccion As String
    QDocs As Long
End Type

Dim CliBien As typCliente
Dim CliMal As typCliente

Dim aDatos As String
Dim aIDCliente As Long, aTipoCliente As Integer
Dim ValAux As Long, gError As Boolean, aResult As Variant

Dim nDirBorrar As Long
Dim idSust As Long, idBorr As Long
Dim BuenoActualiz As Boolean
Dim rBien As rdoResultset, rMal As rdoResultset

Private Sub bBorrar_Click()
    AccionEliminar
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Height = 7245 '5130
    Me.Width = 6740
    
    Picture1.BorderStyle = 0
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    CierroConexion
    
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
    
End Sub


Private Sub CargoDatosCliente(miCliente As Long, miTCliente As Integer, miTextoDatos As String, Cual As typCliente)
    
    Dim miTexto As String
    miTextoDatos = ""
    ClienteACero Cual
    
    If miTCliente = 1 Then
        'cons = "Select Cliente.*, CPersona.*, UAlta.UsuIdentificacion as UsrAlta, UModifica.UsuIdentificacion as UsrModifica" & _
                " From Cliente Left Outer Join Usuario UAlta On CliUsuAlta = UAlta.UsuCodigo" & _
                                   " Left Outer Join Usuario UModifica On CliUsuario = UModifica.UsuCodigo, " & _
                " CPersona " & _
                " Where CliCodigo = CPeCliente And CliCodigo = " & miCliente
        
        Cons = "Select Cliente.*, CPersona.*, UAlta.UsuIdentificacion as UsrAlta, UModifica.UsuIdentificacion as UsrModifica" & _
                " From Cliente Left Outer Join Usuario UAlta On CliUsuAlta = UAlta.UsuCodigo" & _
                                   " Left Outer Join Usuario UModifica On CliUsuario = UModifica.UsuCodigo " & _
                " Left Outer Join CPersona On CPeCliente = CliCodigo" & _
                " Where  CliCodigo = " & miCliente
    End If
    
    If miTCliente = 2 Then
        'cons = "Select Cliente.*, CEmpresa.*, UAlta.UsuIdentificacion as UsrAlta, UModifica.UsuIdentificacion as UsrModifica" & _
                    " From Cliente Left Outer Join Usuario UAlta On CliUsuAlta = UAlta.UsuCodigo" & _
                                   " Left Outer Join Usuario UModifica On CliUsuario = UModifica.UsuCodigo, " & _
                   " CEmpresa " & _
                   " Where CliCodigo = CEmCliente And CliCodigo = " & miCliente
        
        Cons = "Select Cliente.*, CEmpresa.*, UAlta.UsuIdentificacion as UsrAlta, UModifica.UsuIdentificacion as UsrModifica" & _
                    " From Cliente Left Outer Join Usuario UAlta On CliUsuAlta = UAlta.UsuCodigo" & _
                                   " Left Outer Join Usuario UModifica On CliUsuario = UModifica.UsuCodigo " & _
                   " Left Outer Join CEmpresa On CliCodigo = CEmCliente " & _
                   " Where CliCodigo = " & miCliente
    End If
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        
        Cual.Codigo = RsAux!CliCodigo
        If Not IsNull(RsAux!CliDireccion) Then Cual.idDireccion = RsAux!CliDireccion
        If Not IsNull(RsAux!CliCiRuc) Then Cual.CiRuc = RsAux!CliCiRuc
        
        If RsAux!CliTipo = 1 Then    'Si es Persona
            Cual.Tipo = 1
            If Not IsNull(RsAux!CliCiRuc) Then miTextoDatos = "C.I.: " & clsGeneral.RetornoFormatoCedula(Cual.CiRuc) & ", " Else miTextoDatos = "C.I.: N/D, "
             
            miTextoDatos = miTextoDatos & "Alta: " & IIf(Not IsNull(RsAux!CliAlta), Format(RsAux!CliAlta, "dd/mm/yy"), "N/D") & " "
            miTextoDatos = miTextoDatos & "x " & IIf(Not IsNull(RsAux!UsrAlta), Trim(RsAux!UsrAlta), "N/D") & ", "
            
            miTextoDatos = miTextoDatos & "Mod: " & IIf(Not IsNull(RsAux!CliModificacion), Format(RsAux!CliModificacion, "dd/mm/yy"), "N/D") & " "
            miTextoDatos = miTextoDatos & "x " & IIf(Not IsNull(RsAux!UsrModifica), Trim(RsAux!UsrModifica), "N/D")
            miTextoDatos = miTextoDatos & vbCrLf
            If Not IsNull(RsAux!CPeNombre1) Then
                miTexto = Trim(RsAux!CPeNombre1) & " "
                miTexto = miTexto & IIf(Not IsNull(RsAux!CPeNombre2), Trim(RsAux!CPeNombre2) & " ", "")
                miTexto = miTexto & Trim(RsAux!CPeApellido1) & " "
                miTexto = miTexto & IIf(Not IsNull(RsAux!CPeApellido2), Trim(RsAux!CPeApellido2), "")
                miTextoDatos = miTextoDatos & miTexto
                
                Cual.Nombre1 = Trim(RsAux!CPeNombre1)
                If Not IsNull(RsAux!CPeNombre2) Then Cual.Nombre2 = Trim(RsAux!CPeNombre2)
                Cual.Apellido1 = Trim(RsAux!CPeApellido1)
                If Not IsNull(RsAux!CPeApellido2) Then Cual.Apellido2 = Trim(RsAux!CPeApellido2)
                
                miTextoDatos = miTextoDatos & " " & IIf(Not IsNull(RsAux!CPeFNacimiento), "(" & Format(RsAux!CPeFNacimiento, "dd/mm/yy") & ")", "")
                If Not IsNull(RsAux!CPeFNacimiento) Then Cual.FNac = Format(RsAux!CPeFNacimiento, "dd/mm/yy")
            End If
        Else    'Si es Empresa
            Cual.Tipo = 2
            
            miTextoDatos = "RUC: " & IIf(Not IsNull(RsAux!CliCiRuc), clsGeneral.RetornoFormatoRuc(Cual.CiRuc), "N/D") & ", "
            
            miTextoDatos = miTextoDatos & "Alta: " & IIf(Not IsNull(RsAux!CliAlta), Format(RsAux!CliAlta, "dd/mm/yy"), "N/D") & " "
            miTextoDatos = miTextoDatos & "x " & IIf(Not IsNull(RsAux!UsrAlta), Trim(RsAux!UsrAlta), "N/D") & ", "
            
            miTextoDatos = miTextoDatos & "Mod: " & IIf(Not IsNull(RsAux!CliModificacion), Format(RsAux!CliModificacion, "dd/mm/yy"), "N/D") & " "
            miTextoDatos = miTextoDatos & "x " & IIf(Not IsNull(RsAux!UsrModifica), Trim(RsAux!UsrModifica), "N/D")
            miTextoDatos = miTextoDatos & vbCrLf
            
            If Not IsNull(RsAux!CEmFantasia) Then
                Cual.Nombre1 = Trim(RsAux!CEmFantasia)
                miTextoDatos = miTextoDatos & Trim(RsAux!CEmFantasia)
            End If
            If Not IsNull(RsAux!CEmNombre) Then
                Cual.Apellido1 = Trim(RsAux!CEmNombre)
                miTextoDatos = miTextoDatos & " (" & Trim(RsAux!CEmNombre) & ")"
            End If
        End If
        
        Cual.QDocs = QDocumentos(RsAux!CliCodigo)
        miTextoDatos = miTextoDatos & " docs. " & Cual.QDocs
        
        miTextoDatos = miTextoDatos & vbCrLf
        If Not IsNull(RsAux!CliDireccion) Then
            miTexto = clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!CliDireccion)
            miTextoDatos = miTextoDatos & miTexto
            Cual.Direccion = miTexto
        End If
        
    End If
    RsAux.Close
    
End Sub

Private Sub VerificaEstado()
Dim bAzul As Boolean, nMsje As Byte, bDirIgual As Boolean
Dim idSust As Long, idBorr As Long
    
    idSust = Val(tSId.Tag): idBorr = Val(tBId.Tag)
    
    If idSust = 0 Or idBorr = 0 Then
        bAzul = False: nMsje = 1
    Else
        'bDirIgual = IIf(CliBien.FNac = CliMal.FNac, True, False)
        bDirIgual = IIf(CliBien.FNac = CliMal.FNac Or Trim(CliMal.FNac) = "" Or Trim(CliBien.FNac) = "", True, False)
        
        If CliMal.Direccion = CliBien.Direccion And CliMal.Apellido1 Like OrtComodines(CliBien.Apellido1, False) & "*" And _
        bDirIgual And CliMal.Nombre1 Like OrtComodines(CliBien.Nombre1, False) & "*" Then
            bAzul = True: nMsje = 3
        Else
            bAzul = False: nMsje = 2
        End If
    End If
    
    lEstado.BackColor = IIf(bAzul, 10485760, 26367)
    Select Case nMsje
        Case 1: lEstado.Caption = "Debe completar los datos"
        Case 2: lEstado.Caption = "Datos posiblemente incorrectos"
        Case 3: lEstado.Caption = "Listo para sustituir"
    End Select
        
    If nMsje <> 1 And CliBien.CiRuc <> "" And CliMal.CiRuc <> "" Then lEstado.Caption = "OJO!!! Ambos clientes tienen Cédula!"
    
    If idBorr = 0 Or idSust = 0 Then bBorrar.Enabled = False Else bBorrar.Enabled = True
    
End Sub

Private Function QDocumentos(miCliente As Long) As Long
    On Error GoTo errQDoc
    Dim rsDoc As rdoResultset
    QDocumentos = 0
    
    Cons = "Select Count(*) From Documento Where DocCliente = " & miCliente
    Set rsDoc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsDoc.EOF Then If Not IsNull(rsDoc(0)) Then QDocumentos = rsDoc(0)
    rsDoc.Close
    
errQDoc:
End Function

Private Sub AccionEliminar()

    Dim bHay As Boolean
    
    If Val(tBId.Tag) = 0 Or Val(tSId.Tag) = 0 Then
        MsgBox "Falta seleccionar alguno de los clientes.", vbExclamation, "Faltan Clientes"
        Exit Sub
    End If
    
    If Val(tBId.Tag) = Val(tSId.Tag) Then
        MsgBox "Los clientes seleccionados son los mismos.", vbExclamation, "Faltan Clientes"
        Exit Sub
    End If
    
    'Ver si es cambio de Persona a Empresa
    If CliMal.Tipo = 2 And CliBien.Tipo = 1 Then
        
        Screen.MousePointer = 11
        bHay = False
        Cons = "Select * from EmpresaDato Where EDaCodigo=" & CliMal.Codigo
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then bHay = True
        RsAux.Close
        
        Screen.MousePointer = 0
        If bHay Then
            MsgBox "No se puede sustituir a un Proveedor nuestro por un cliente Persona!", vbInformation, "Posible Error"
            Exit Sub
        End If
    End If
    
    If CliMal.QDocs > CliBien.QDocs Then
        If MsgBox("El cliente 'Malo' tiene mas operaciones que el cliente sustituto." & vbCrLf & _
                       "Quiere verificar los datos (se sugiere invertir los clientes)", vbQuestion + vbYesNo, "Verificar") = vbYes Then Exit Sub
    End If
    
    
    If MsgBox("¿Está seguro que quiere Traspasar las Operaciones del Cliente?", vbQuestion + vbYesNo, "Confirma") = vbNo Then Exit Sub
    On Error GoTo FinPorError
    
    idBorr = CliMal.Codigo: idSust = CliBien.Codigo
    tReporte.Text = ""
    
    sBar.Panels(1).Text = " Sustituyendo Cliente, por favor espere..."
    tab1.SelectedItem = tab1.Tabs("proceso")
    Me.Refresh
    Screen.MousePointer = 11
    
    'Elimina en Envios.
    'LlenaReporte sql_Update("Envio", "EnvCliente", idBorr, idSust)
    LlenaReporte sql_Update("Envio", "EnvCliente", idBorr, idSust, _
                        sqlAnd:=" EnvTipo In (2,3) And EnvDocumento In (Select VTeCodigo from VentaTelefonica Where VTeCliente = " & idBorr & ")")
    
    LlenaReporte sql_Update("Envio", "EnvCliente", idBorr, idSust, _
                        sqlAnd:=" EnvTipo = 1 And EnvDocumento In (Select DocCodigo from Documento Where DocCliente =  " & idBorr & ")")
    
    
    'Sustituye en la tabla Credito y Documento, Vtas.Telef, Colectivos y Service (prod).
    LlenaReporte sql_Update("Documento", "DocCliente", idBorr, idSust)
    LlenaReporte sql_Update("Credito", "CreCliente", idBorr, idSust)
    LlenaReporte sql_Update("Credito", "CreGarantia", idBorr, idSust)
    LlenaReporte sql_Update("VentaTelefonica", "VTeCliente", idBorr, idSust)
    LlenaReporte sql_Update("Colectivo", "ColCliente1", idBorr, idSust)
    LlenaReporte sql_Update("Colectivo", "ColCliente2", idBorr, idSust)
    LlenaReporte sql_Update("CuentaDocumento", "CDoIDTipo", idBorr, idSust, , "CDoTipo = 1")
    LlenaReporte sql_Update("Producto", "ProCliente", idBorr, idSust)
    LlenaReporte sql_Update("Servicio", "SerCliente", idBorr, idSust)

    'Sustituye en las Tablas del Clearing.
    'LlenaReporte sql_Update("Clearing", "CleCliente", idBorr, idSust)
    LlenaReporte ComparoClearing
    LlenaReporte sql_Update("ClearingAntecedente", "CAnCliente", idBorr, idSust)
    LlenaReporte sql_Update("ClearingCheque", "CChCliente", idBorr, idSust)
    LlenaReporte sql_Update("ClearingSolicitud", "CSoCliente", idBorr, idSust)

    'Sustituye Cliente en la tabla Solicitud.
    LlenaReporte sql_Update("Solicitud", "SolCliente", idBorr, idSust)
    LlenaReporte sql_Update("Solicitud", "SolGarantia", idBorr, idSust)
    
    'Elimina en Envios (*) antes estaba aca.
    
    'Sustituye Llamadas de Clientes
    LlenaReporte sql_Update("Llamada", "LlaCliente", idBorr, idSust)
    LlenaReporte sql_Update("TelefonoExtra", "TExCliente", idBorr, idSust)
    

    LlenaReporte sql_Update("Suceso", "SucCliente", idBorr, idSust)
    
    LlenaReporte sql_Update("Devolucion", "DevCliente", idBorr, idSust)
    LlenaReporte sql_Update("ClienteWeb", "CWeCliente", idBorr, idSust)
    LlenaReporte sql_Update("DireccionesBorradas", "DBoCliente", idBorr, idSust)
    LlenaReporte sql_Update("TelefonosBorrados", "TBoCliente", idBorr, idSust)
    
    'Tabla Mensaje Links de la logdb    11/05/2004
    LlenaReporte sql_Update_OLD("logdb.dbo.MensajeLinks", "MLiRegistro", idBorr, idSust, sqlAnd:="MLiTabla  IN (Select TabCodigo from logdb.dbo.Tablas Where TabNombre = 'CGSA.dbo.Cliente' And TabCampoIdId = 'CliCodigo')")
    
' *** A partir de acá Se comienzan a eliminar datos del Cliente
    'If Me!Eliminar = False Then GoTo FinProceder
    'Elimina Comentarios del Cliente.
    LlenaReporte sql_Update("Comentario", "ComCliente", idBorr, idSust)
    'Traspasa Direcciones auxiliares
    LlenaReporte sql_Update("DireccionAuxiliar", "DAuCliente", idBorr, idSust)

    'Traspasa Mails
    LlenaReporte sql_Update("EMailDireccion", "EMDIdCliente", idBorr, idSust)
    
    'Elimina Relaciones del Cliente.
    LlenaReporte sql_Delete("PersonaRelacion", "PReClienteEs", idBorr, " PReClienteDe = " & idSust) 'Borro Rel Entre Ellos
    LlenaReporte sql_Delete("PersonaRelacion", "PReClienteDe", idBorr, " PReClienteEs = " & idSust)
    
    'Borro Rel que van a duplicarse
    LlenaReporte sql_Delete("PersonaRelacion", "PReClienteEs", idBorr, " PReClienteDe IN (Select PReClienteDe FRom PersonaRelacion Where PReClienteEs = " & idSust & ")")
    LlenaReporte sql_Delete("PersonaRelacion", "PReClienteDe", idBorr, " PReClienteEs IN (Select PReClienteEs FRom PersonaRelacion Where PReClienteDe = " & idSust & ")")
    
    LlenaReporte sql_Update("PersonaRelacion", "PReClienteEs", idBorr, idSust)
    LlenaReporte sql_Update("PersonaRelacion", "PReClienteDe", idBorr, idSust)
    
    'Elimina Títulos del Cliente.
    LlenaReporte sql_Update("Titulo", "TitCliente", idBorr, idSust)
    
    'Elimina Teléfonos del Cliente.
    'LlenaReporte sql_Update("Telefono", "TelCliente", idBorr, idSust, gError)
    Cons = sql_Update("Telefono", "TelCliente", idBorr, idSust, gError)
    If gError Then
        ComparoTelefonos idBorr, idSust
    Else
        LlenaReporte Cons
    End If
    
    'Elimina Referencias del Cliente.
    LlenaReporte sql_Update("ReferenciaCliente", "RClCPersona", idBorr, idSust)
    
    'Elimina Empleos del Cliente.
    LlenaReporte sql_Update("Empleo", "EmpCliente", idBorr, idSust)
    
'22/3/2011 Desde que se deja ingresar personas como empleadoras.
    'If CliMal.Tipo = 2 Then
    LlenaReporte sql_Update("Empleo", "EmpCEmpresa", idBorr, idSust)
    
    'Elimina Persona o Empresa del Cliente.
    ComparaClientes

    '22/3/2011 nuevos casos --------------------------------------------------
    
    'LlenaReporte sql_Update("AuxEstadisticasCreditos", "AECCliente", idBorr, idSust)
    
    LlenaReporte sql_Update("ChequeDiferido", "CDiCliente", idBorr, idSust)
    LlenaReporte sql_Update("ChequeDiferido", "CDiClienteFactura", idBorr, idSust)
    LlenaReporte sql_Update("CambioArticulo", "CArDCliente", idBorr, idSust)
    LlenaReporte sql_Update("CambioArticulo", "CArSIDCliente", idBorr, idSust)
    LlenaReporte sql_Update("CartaMoroso", "CMoCliente", idBorr, idSust)
    LlenaReporte sql_Update("ClearingComentario", "CCoCliente", idBorr, idSust)
    LlenaReporte sql_Update("ClearingMasivo", "CMaCliente", idBorr, idSust)
    LlenaReporte sql_Update("ClientePlazo", "CPlCliente", idBorr, idSust)
    LlenaReporte sql_Update("ComTransacciones", "TraClienteCGSA", idBorr, idSust)
    LlenaReporte sql_Update("CotizacionCliente", "CClCliente", idBorr, idSust)
    LlenaReporte sql_Update("CreditoPreAutorizado", "CPACliente", idBorr, idSust)
    LlenaReporte sql_Update("DocumentosASacar", "DASCliente", idBorr, idSust)
    LlenaReporte sql_Update("Quotation", "QuoCliente", idBorr, idSust)
    
' no hay permisos y da lio con la contrainst.  LlenaReporte sql_Update("ClienteWebUsuario", "TraClienteCGSA", idBorr, idSust)
    
    '22/3/2011 nuevos casos --------------------------------------------------

    If CliMal.Tipo = 1 Then 'Si es Persona
        ComparaCPersonas
        LlenaReporte sql_Delete("CPersona", "CPeCliente", idBorr)
        
            
    Else    'Si es Empresa
        LlenaReporte sql_Update("SucursalDeBanco", "SBaBanco", idBorr, idSust)
        
        LlenaReporte sql_Update("ChequeDiferido", "CDiBanco", idBorr, idSust)
        
        LlenaReporte sql_Update("Referencia", "RefEmpresa", idBorr, idSust)
        ComparaCEmpresas
        LlenaReporte sql_Delete("CEmpresa", "CEmCliente", idBorr)
        
        'Empresa Dato ----------------------------------------------------------------------------------------------------------------
        bHay = False
        Cons = "Select * from EmpresaDato Where EDaCodigo = " & idSust
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then bHay = True
        RsAux.Close
        
        If bHay Then   'Si el bueno YA es Proveedor Nuestro.
            LlenaReporte sql_Delete("EmpresaDato", "EDaCodigo", idBorr)  'Borro al malo como proveedor
        Else
            LlenaReporte sql_Update("EmpresaDato", "EDaCodigo", idBorr, idSust)  'Sustituyo al malo x el bueno
        End If
        '----------------------------------------------------------------------------------------------------------------------------------
        
    'Sustituye Cliente de la Tabla "Compra"
'22/3/2011 Eliminamos estos 2 update ya que estos campos ahora están en la genEntes en BD Zureo.
'        LlenaReporte sql_Update("Compra", "ComProveedor", idBorr, idSust)
'        LlenaReporte sql_Update("RemitoCompra", "RCoProveedor", idBorr, idSust)
        
        'Importaciones
        LlenaReporte sql_Update("Carpeta", "CarBcoEmisor", idBorr, idSust)
        LlenaReporte sql_Update("Embarque", "EmbAgencia", idBorr, idSust)
        
    End If
    
    ValAux = InStr(1, tReporte.Text, "Error")
    If ValAux > 0 Then
        GoTo FinPorError
        Exit Sub
    End If
    
    'Elimina Cliente de la tabla Cliente.
    LlenaReporte sql_Delete("Cliente", "CliCodigo", idBorr)
    
    If nDirBorrar > 0 Then
        Cons = "Select * from Cliente Where CliCodigo=" & idSust
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then ValAux = RsAux!CliDireccion
        RsAux.Close
        
        'If CliMal.Tipo = 2 Then
        LlenaReporte sql_Update("Empleo", "EmpDirEmpleo", nDirBorrar, ValAux)
        
        LlenaReporte sql_Delete("Direccion", "DirCodigo", nDirBorrar)
    End If
    
    'Hago el cambio en la bd de la web.
'    Cons = "SET ANSI_NULLS ON SET ANSI_WARNINGS ON update WebServer.WebCG.dbo.cliente set cliclicodigo = " & idSust & " where cliid = " & idBorr
'    cBase.Execute Cons
    
FinProceder:
    GrabaEnCliBorrados
    GrabaComentario idSust, "Absorbió Cliente duplicado Id: " & idBorr & " " & CliMal.CiRuc & ", " & Trim(Trim(CliMal.Nombre1) & " " & Trim(CliMal.Nombre2)) & " " & Trim(Trim(CliMal.Apellido1) & " " & Trim(CliMal.Apellido2)), _
                                paCodigoDeUsuario
    tReporte.FontSize = 8
    
    ValAux = InStr(1, tReporte.Text, "Error")
    If ValAux > 0 Then
        MsgBox Mid$(tReporte.Text, ValAux), vbCritical, "Atención (hubo errores)"
    Else
        MsgBox "Tarea finalizada exitosamente", vbInformation, "Buen Fin"
    End If
    
    lEstado.BackColor = 26367: lEstado.Caption = "Completar los datos"
    tBId.Text = "": tSId.Text = "": tBCiRuc.Text = "": tSCiRuc.Text = ""
    tBTexto.Text = "": tSTexto.Text = ""
    tSCiRuc.SetFocus
    bBorrar.Enabled = False
    
    sBar.Panels(1).Text = ""
    Screen.MousePointer = 0
    Exit Sub
    
FinPorError:
    GrabaEnCliBorrados
    tReporte.FontSize = 8
    
    MsgBox Mid$(tReporte.Text, ValAux), vbCritical, "Atención (hubo errores)"
    
    sBar.Panels(1).Text = "ERROR"
    Screen.MousePointer = 0
    
End Sub

Private Sub ComparaCPersonas()
    'Casa al conyuge con el Cliente nuevo
    'El mal es persona ---> el bien no se sabe
    'CliBien.Tipo
    
    On Error Resume Next
    LlenaReporte sql_Update("CPersona", "CPeConyuge", idBorr, idSust)
    
    Cons = "SELECT * FROM CPersona Where CPeCliente=" & idBorr
    Set rMal = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Cons = "SELECT * FROM CPersona Where CPeCliente=" & idSust
    Set rBien = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If CliBien.Tipo = 2 Then        ' el bien es una empresa
        Dim aValor As Long: aValor = 0
        If Not IsNull(rMal!CPeConyuge) Then aValor = rMal!CPeConyuge
        rBien.Close: rMal.Close
        
        GrabaComentario aValor, "Cónyuge Borrado x Sustituir_Cliente. Se paso a id_Cliente: " & idSust & " " & _
                                               Trim(CliBien.Nombre1) & " " & Trim(CliBien.Apellido1) & " (id_Old: " & idBorr & ")", paCodigoDeUsuario
                                               
        GrabaComentario idSust, "Cónyuge Borrado x Sustituir_Cliente. Tenía como Cónyuge a id_Cliente: " & aValor & _
                                               " se pierde referencia al sustituir.", paCodigoDeUsuario
        Exit Sub
    End If
    
    
    rBien.Edit
    If Not IsNull(rMal!CPeFNacimiento) Then
        If Not IsNull(rBien!CPeFNacimiento) Then
            If CDate(rBien!CPeFNacimiento) < CDate("1/1/1900") Then rBien!CPeFNacimiento = rMal!CPeFNacimiento
        Else
            rBien!CPeFNacimiento = rMal!CPeFNacimiento
        End If
    End If
    If IsNull(rBien!CPeNombre2) Then rBien!CPeNombre2 = rMal!CPeNombre2
    If IsNull(rBien!CPeApellido2) Then rBien!CPeApellido2 = rMal!CPeApellido2
    ComparoDatos "CPePropietario"
    ComparoDatos "CPeRuc"
    ComparoDatos "CPeConyuge"
    ComparoDatos "CPeOcupacion"
    ComparoDatos "CPeEstadoCivil"
    ComparoDatos "CPeSexo"
    
    rBien.Update: rMal.Close: rBien.Close

End Sub

Private Sub ComparaCEmpresas()
    On Error Resume Next
    'El malo es empresa ---> el bueno ??
    If CliBien.Tipo = 1 Then Exit Sub
    
    Cons = "SELECT * FROM CEmpresa Where CEmCliente=" & idBorr
    Set rMal = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Cons = "SELECT * FROM CEmpresa Where CEmCliente=" & idSust
    Set rBien = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    rBien.Edit
    ComparoDatos "CEmRamo"
    ComparoDatos "CEmAfiliado"
    ComparoDatos "CEmEstatal"
    rBien.Update: rMal.Close: rBien.Close
    
End Sub

Private Sub ComparaClientes()

    Cons = "SELECT * FROM Cliente WHERE CliCodigo=" & idBorr
    Set rMal = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Cons = "SELECT * FROM Cliente WHERE CliCodigo=" & idSust
    Set rBien = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    rBien.Edit
    If rMal!CliModificacion > rBien!CliModificacion Then
        BuenoActualiz = False
        rBien!CliModificacion = rMal!CliModificacion
    Else
        BuenoActualiz = True
    End If
    
    ComparoDatos "CliUsuario"

    If rMal!CliAlta < rBien!CliAlta Then rBien!CliAlta = rMal!CliAlta
    If rMal!CliCheque = "S" Or rBien!CliCheque = "S" Then rBien!CliCheque = "S"

    If Not IsNull(rMal!CliSolicitud) Then
        If IsNull(rBien!CliSolicitud) Then
            rBien!CliSolicitud = rMal!CliSolicitud
        Else
            If rMal!CliSolicitud <> rBien!CliSolicitud Then
                GrabaComentario idSust, "Solicitud anterior " & rMal!CliSolicitud, miConexion.UsuarioLogueado(True, False)
    End If: End If: End If
    
    If Not IsNull(rBien!CliDireccion) Then ValAux = rBien!CliDireccion Else ValAux = 0
    If Not IsNull(rMal!CliDireccion) Then nDirBorrar = rMal!CliDireccion Else nDirBorrar = 0
    ComparoDatos "CliDireccion"
    
    If nDirBorrar = rBien!CliDireccion Then
        nDirBorrar = ValAux
        If Not IsNull(rBien!CliDireccion) Then ValAux = rBien!CliDireccion Else ValAux = 0
    End If

    ComparoDatos "CliEMail"

    If IsNull(rBien!CliCiRuc) Then rBien!CliCiRuc = Trim(rMal!CliCiRuc)
    rBien.Update: rMal.Close: rBien.Close
    
    If nDirBorrar > 0 Then ComparoDireccion ValAux, nDirBorrar
    
End Sub

Private Sub ComparoDireccion(Buena As Long, Mala As Long)
    
    'Veo si las 2 direcc son iguales.
    Cons = "SELECT * FROM Direccion WHERE DirCodigo=" & Mala
    Set rMal = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Cons = "SELECT * FROM Direccion WHERE DirCodigo=" & Buena
    Set rBien = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If (rBien!DirPuerta <> 0 And rBien!DirPuerta = rMal!DirPuerta) Or _
    (rBien!DirCalle = rMal!DirCalle And rBien!DirPuerta = rMal!DirPuerta) Then
        nDirBorrar = Mala
        rBien.Edit
        ComparoDatos "DirLetra": ComparoDatos "DirBis"
        ComparoDatos "DirApartamento": ComparoDatos "DirSenda"
        ComparoDatos "DirBloque": ComparoDatos "DirEntre1"
        ComparoDatos "DirEntre2": ComparoDatos "DirAmpliacion"
        ComparoDatos "DirConfirmada": ComparoDatos "DirVive"
        ComparoDatos "DirCampo1": ComparoDatos "DirCampo2": ComparoDatos "DirComplejo"
        rBien.Update
    Else
        'Se agrega como Dirección auxiliar.
        Cons = "Select * from DireccionAuxiliar Where DAuCliente = " & idSust & " And DAuDireccion = " & nDirBorrar
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux.EOF Then
            RsAux.AddNew
            RsAux!DAuCliente = idSust
            RsAux!DAuDireccion = nDirBorrar
            RsAux!DAuNombre = "de Cliente duplicado"
            RsAux.Update
        End If
        RsAux.Close
            
        'cons = "INSERT INTO DireccionAuxiliar (DAuCliente, DAuDireccion, DAuNombre) SELECT " & idSust & ", " & nDirBorrar & ", 'de Cliente duplicado'"
        'cBase.Execute cons
        nDirBorrar = 0  'No hay q borrar xq es otra dirección.
    End If
    
    rMal.Close: rBien.Close
    
End Sub
    
Private Sub ComparoDatos(sCampo As String)
    
    If rBien(sCampo) = rMal(sCampo) Then Exit Sub
    If Not IsNull(rMal(sCampo)) Then
        If IsNull(rBien(sCampo)) Or Not BuenoActualiz Then
            rBien(sCampo) = rMal(sCampo)
        End If
    End If
    
    'If cBien = cMal Then Exit Function
    'If Not IsNull(cMal) Then
    '    If IsNull(cBien) Or Not BuenoActualiz Then ComparoDatos = True
    'End If

End Sub

Private Sub ComparoTelefonos(idBorr As Long, idSust As Long)
    Dim rsTel As rdoResultset
    Dim bExiste As Boolean
    
    Cons = "SELECT * FROM Telefono WHERE TelCliente=" & idBorr
    Set rsTel = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do Until rsTel.EOF
        bExiste = False
        Cons = "Select * from Telefono Where TelCliente = " & idSust & " And TelNumero='" & Trim(rsTel!TelNumero) & "'"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then bExiste = True
        RsAux.Close
        
        If bExiste Then
            rsTel.Delete 'Lo borro xq Existe
        Else
            ValAux = ProxTipoDisponible(idSust, rsTel!TelTipo)
            If ValAux > 0 Then
                
                rsTel.Edit
                rsTel!TelTipo = ValAux
                rsTel!TelCliente = idSust
                If IsNull(rsTel!telInterno) Then rsTel!telInterno = "Migrado"
                
                LlenaReporte "UpD. 1 Telefono. TelCliente " & "=" & idBorr & " TelNumero = " & rsTel!TelNumero
                rsTel.Update
                
            Else    'No hay un Tipo Disponible.
                LlenaReporte "Borr. Teléf. (no hay tipo): (Tipo: " & rsTel!TelTipo & ") " & Trim(rsTel!TelNumero)
                rsTel.Delete
            End If
        End If
        rsTel.MoveNext
    Loop
    rsTel.Close
End Sub

Private Function ComparoClearing()
Dim rsCC As rdoResultset

    'LlenaReporte sql_Update("Clearing", "CleCliente", idBorr, idSust)
    On Error GoTo errComparoClearing
    Dim bcMalo As Boolean, bcBueno As Boolean
    Dim scMalo As String, scBueno As String
    
    bcMalo = False: bcBueno = True
    
    'Valido si el cliente malo tiene datos en la tabla clearing
    Cons = "Select * From Clearing Where CleCliente = " & idBorr
    Set rsCC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsCC.EOF Then
        bcMalo = True
        If Not IsNull(rsCC!CleNombre) Then scMalo = Trim(rsCC!CleNombre)
        If Not IsNull(rsCC!CleApellidos) Then scMalo = scMalo & Trim(rsCC!CleApellidos)
    End If
    rsCC.Close
    
    If Not bcMalo Then
        ComparoClearing = sql_Update("Clearing", "CleCliente", idBorr, idSust)
        Exit Function
    End If
    
    'Valido si el cliente Bueno tiene datos en la tabla clearing
    Cons = "Select * From Clearing Where CleCliente = " & idSust
    Set rsCC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsCC.EOF Then
        bcBueno = True
        If Not IsNull(rsCC!CleNombre) Then scBueno = Trim(rsCC!CleNombre)
        If Not IsNull(rsCC!CleApellidos) Then scBueno = scBueno & Trim(rsCC!CleApellidos)
    End If
    rsCC.Close
    
    If Not bcBueno Then
        ComparoClearing = sql_Update("Clearing", "CleCliente", idBorr, idSust)
        Exit Function
    End If
    
    'Los dos clientes tiene datos en la tabla Clearing
    If Trim(scMalo) = "??" Then
        bcMalo = False
    Else
        If Trim(scBueno) = "??" Then
            bcBueno = False
        Else
            bcMalo = False
        End If
    End If
    
    If bcBueno Then
        'Si me quedo con el bueno borro el malo
        ComparoClearing = sql_Delete("Clearing", "CleCliente", idBorr)
    Else
        If bcMalo Then
            Dim sRet As String
            sRet = sql_Delete("Clearing", "CleCliente", idSust)
            sRet = sRet & vbCrLf
            sRet = sRet & sql_Update("Clearing", "CleCliente", idBorr, idSust)
            ComparoClearing = sRet
        End If
    End If
    Exit Function

errComparoClearing:
    ComparoClearing = "Error ComparoClearing " & Err.Description
End Function

Private Sub tab1_Click()
    Select Case LCase(tab1.SelectedItem.Key)
        Case "datos": Picture1.ZOrder 0
        Case "proceso": tReporte.ZOrder 0
    End Select
End Sub

Private Sub tBCiRuc_Change()
    tBId.Tag = 0
    tBTexto.Text = ""
End Sub

Private Sub tBCiRuc_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF12: EjecutarApp App.Path & "\Visualizacion de Operaciones " & Val(tBId.Tag)
        Case vbKeyF4, vbKeyF5: BuscoCliente KeyCode, ABorrar:=True
    End Select
        
End Sub

Private Sub LlenaReporte(Txto As String)
    tReporte.Text = tReporte.Text & Txto & vbCrLf
    tReporte.Refresh
End Sub

Private Sub tBCiRuc_KeyPress(KeyAscii As Integer)
On Error GoTo errBuscar
    
    If KeyAscii = vbKeyReturn Then
        If Val(tBId.Tag) <> 0 Then
            If Val(tSId.Tag) <> 0 Then bBorrar.SetFocus Else tSId.SetFocus
            Exit Sub
        End If
        tBCiRuc.Text = Replace(Replace(Replace(tBCiRuc, ".", ""), "-", ""), " ", "")
        If Not IsNumeric(tBCiRuc.Text) Then Exit Sub
        
        Screen.MousePointer = 11
        aIDCliente = 0
        
        Cons = "Select * from Cliente Where CliCIRuc = '" & Trim(tBCiRuc.Text) & "'"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            aTipoCliente = RsAux!CliTipo
            aIDCliente = RsAux!CliCodigo
        End If
        RsAux.Close
        
        If aIDCliente > 0 Then
            CargoDatosCliente aIDCliente, aTipoCliente, aDatos, CliMal
            tBId.Text = CliMal.Codigo
            tBCiRuc.Text = CliMal.CiRuc
            tBTexto.Text = aDatos
            
            tBId.Tag = CliMal.Codigo
        End If
        
        VerificaEstado
        
        If Val(tBId.Tag) <> 0 Then
            If Val(tSId.Tag) <> 0 Then bBorrar.SetFocus Else tSId.SetFocus
        End If
        Screen.MousePointer = 0
    End If
    Exit Sub
    
errBuscar:
    'MsgBox Err.Number & Err.Description
    clsGeneral.OcurrioError "Error al buscar el cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tBId_Change()
    tBId.Tag = 0
    tBTexto.Text = ""
End Sub

Private Sub tBId_KeyPress(KeyAscii As Integer)
    
    On Error GoTo errBuscar
    
    If KeyAscii = vbKeyReturn Then
        If Val(tBId.Tag) <> 0 Then tBCiRuc.SetFocus: Exit Sub
        If Trim(tBId.Text) = "" Then tBCiRuc.SetFocus: Exit Sub
        If Not IsNumeric(tBId.Text) Then Exit Sub
        
        Screen.MousePointer = 11
        aIDCliente = Val(tBId.Text)
        
        Cons = "Select * from Cliente Where CliCodigo = " & aIDCliente
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then aTipoCliente = RsAux!CliTipo Else aIDCliente = 0
        RsAux.Close
        
        If aIDCliente > 0 Then
            CargoDatosCliente aIDCliente, aTipoCliente, aDatos, CliMal
            tBId.Text = CliMal.Codigo
            tBCiRuc.Text = CliMal.CiRuc
            tBTexto.Text = aDatos
            
            tBId.Tag = CliMal.Codigo
        End If
        VerificaEstado
        Screen.MousePointer = 0
    End If
    Exit Sub
    
errBuscar:
    clsGeneral.OcurrioError "Error al buscar el cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub ClienteACero(ID As typCliente)
    ID.Apellido1 = ""
    ID.Apellido2 = ""
    ID.CiRuc = ""
    ID.Codigo = 0
    ID.FNac = ""
    ID.idDireccion = 0
    ID.Nombre1 = ""
    ID.Nombre2 = ""
    ID.QDocs = 0
    ID.Tipo = 0
    ID.Direccion = ""
End Sub

Private Sub tSCiRuc_Change()
    tSId.Tag = 0
    tSTexto.Text = ""
End Sub

Private Sub tSCiRuc_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF12: EjecutarApp App.Path & "\Visualizacion de Operaciones " & Val(tSId.Tag)
        Case vbKeyF4, vbKeyF5: BuscoCliente KeyCode, ABorrar:=False
    End Select
End Sub

Private Sub tSCiRuc_KeyPress(KeyAscii As Integer)
On Error GoTo errBuscar
    
    If KeyAscii = vbKeyReturn Then
        If Val(tSId.Tag) <> 0 Then tBCiRuc.SetFocus: Exit Sub
        
        tSCiRuc.Text = Replace(Replace(Replace(tSCiRuc, ".", ""), "-", ""), " ", "")
        
        If Not IsNumeric(tSCiRuc.Text) Then Exit Sub
        
        Screen.MousePointer = 11
        aIDCliente = 0
        
        Cons = "Select * from Cliente Where CliCIRuc = '" & Trim(tSCiRuc.Text) & "'"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            aTipoCliente = RsAux!CliTipo
            aIDCliente = RsAux!CliCodigo
        End If
        RsAux.Close
        
        If aIDCliente > 0 Then
            CargoDatosCliente aIDCliente, aTipoCliente, aDatos, CliBien
            tSId.Text = CliBien.Codigo
            tSCiRuc.Text = CliBien.CiRuc
            tSTexto.Text = aDatos
            
            tSId.Tag = CliBien.Codigo
            
            aIDCliente = BuscoClienteSustituto
            
            If aIDCliente <> 0 Then
                tBId.Text = aIDCliente
                Call tBId_KeyPress(vbKeyReturn)
            End If
            
        End If
        VerificaEstado
        If Val(tSId.Tag) <> 0 Then tBCiRuc.SetFocus
        Screen.MousePointer = 0
    End If
    Exit Sub
    
errBuscar:
    clsGeneral.OcurrioError "Error al buscar el cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tSId_Change()
    tSId.Tag = 0
    tSTexto.Text = ""
End Sub

Private Sub tSId_KeyPress(KeyAscii As Integer)
On Error GoTo errBuscar
    
    If KeyAscii = vbKeyReturn Then
        If Val(tSId.Tag) <> 0 Then tSCiRuc.SetFocus: Exit Sub
        If Trim(tSId.Text) = "" Then tSCiRuc.SetFocus: Exit Sub
        If Not IsNumeric(tSId.Text) Then Exit Sub
        
        Screen.MousePointer = 11
        aIDCliente = Val(tSId.Text)
        
        Cons = "Select * from Cliente Where CliCodigo = " & aIDCliente
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then aTipoCliente = RsAux!CliTipo Else aIDCliente = 0
        RsAux.Close
        
        If aIDCliente > 0 Then
            CargoDatosCliente aIDCliente, aTipoCliente, aDatos, CliBien
            tSId.Text = CliBien.Codigo
            tSCiRuc.Text = CliBien.CiRuc
            tSTexto.Text = aDatos
            
            tSId.Tag = CliBien.Codigo
            
            aIDCliente = BuscoClienteSustituto
            
            If aIDCliente <> 0 Then
                tBId.Text = aIDCliente
                Call tBId_KeyPress(vbKeyReturn)
            End If
            
        End If
        VerificaEstado
        Screen.MousePointer = 0
    End If
    Exit Sub
    
errBuscar:
    clsGeneral.OcurrioError "Error al buscar el cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub BuscoCliente(KeyCode As Integer, Optional ABorrar As Boolean = True)
    On Error GoTo errBuscar
    
    If KeyCode = vbKeyF4 Or KeyCode = vbKeyF5 Then
        Screen.MousePointer = 11
        Dim miObj As New clsBuscarCliente
        miObj.ActivoFormularioBuscarClientes cBase, KeyCode = vbKeyF4, KeyCode = vbKeyF5
        aIDCliente = miObj.BCClienteSeleccionado
        Set miObj = Nothing
        Me.Refresh
        
        If aIDCliente <> 0 Then
            If ABorrar Then tBId.Text = aIDCliente: Call tBId_KeyPress(vbKeyReturn)
            If Not ABorrar Then tSId.Text = aIDCliente: Call tSId_KeyPress(vbKeyReturn)
        End If
        Screen.MousePointer = 0
    End If
    Exit Sub

errBuscar:
    clsGeneral.OcurrioError "Error al buscar el cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function BuscoClienteSustituto() As Long
    On Error GoTo errBuscar
    'Intenta buscar el cliente que está duplicado (si es que no fue ingresado)
    aIDCliente = 0
    BuscoClienteSustituto = aIDCliente
    If Val(tBId.Tag) <> 0 Then Exit Function

    Cons = "SELECT * FROM Cliente, CPersona  " & _
                " Where CliCodigo = CPeCliente " & _
                " And CliCodigo <>" & Val(tSId.Tag) & _
                " AND CPeNombre1 = '" & Trim(CliBien.Nombre1) & "'" & _
                " AND CPeApellido1 = '" & Trim(CliBien.Apellido1) & "'"
    
'    If Trim(CliBien.Nombre2) <> "" Then cons = cons & " And (CPeNombre2 = '" & Trim(CliBien.Nombre2) & "' OR CPeNombre2 is null) "
'    If Trim(CliBien.Apellido2) <> "" Then cons = cons & " And (CPeApellido2 = '" & Trim(CliBien.Apellido2) & "' OR CPeApellido2 is null) "
        
    If Trim(CliBien.Nombre2) <> "" Then Cons = Cons & " And CPeNombre2 = '" & Trim(CliBien.Nombre2) & "'"
    If Trim(CliBien.Apellido2) <> "" Then Cons = Cons & " And CPeApellido2 = '" & Trim(CliBien.Apellido2) & "'"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        aIDCliente = RsAux!CliCodigo
        RsAux.MoveNext
        If Not RsAux.EOF Then aIDCliente = 0
    End If
    RsAux.Close
    
    BuscoClienteSustituto = aIDCliente
    Exit Function

errBuscar:
    clsGeneral.OcurrioError "Error al buscar el sustituto.", Err.Description
End Function

Private Sub GrabaEnCliBorrados()
    On Error GoTo errGCB
    'Graba el Registro en la tabla Clientes Borrados
    'CBoCliente  CBoFecha                    CBoSustituto CBoReporte
    Dim aFecha As Date
    
    aFecha = Now
    
    Cons = "Select CBoCliente, CBoFecha, CBoUsuario, CBoSustituto from ClientesBorrados Where CBoCliente = " & idBorr
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.AddNew
    RsAux!CBoFecha = Format(aFecha, "mm/dd/yyyy hh:mm:ss")
    RsAux!CBoCliente = idBorr: RsAux!CBoSustituto = idSust
    'rsAux!CBoReporte = Trim(tReporte.Text)
    RsAux!CBoUsuario = paCodigoDeUsuario
    
    RsAux.Update: RsAux.Close
    
    Cons = "Update ClientesBorrados Set CBoReporte = '" & Trim(tReporte.Text) & "'" & _
              " Where CBoCliente = " & idBorr & _
              " And CBoSustituto = " & idSust & _
              " And CBoFecha = '" & Format(aFecha, "mm/dd/yyyy hh:mm:ss") & "'"

    cBase.Execute Cons
    Exit Sub

errGCB:
    clsGeneral.OcurrioError "Error al grabar en ClientesBorrados.", Err.Description
End Sub

Private Sub GrabaComentario(idCliente As Long, sComent As String, Usuario As Long)
    On Error GoTo errGC
    Dim rsCom As rdoResultset
    
    Cons = "Select * from Comentario Where ComCliente = " & idCliente
    Set rsCom = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    rsCom.AddNew
    rsCom!ComCliente = idCliente
    rsCom!ComFecha = Now
    rsCom!ComComentario = Trim(sComent)
    rsCom!ComTipo = 1
    rsCom!ComUsuario = Usuario
    rsCom.Update: rsCom.Close
    Exit Sub

errGC:
    clsGeneral.OcurrioError "Error al grabar comentarios.", Err.Description
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With tab1
        .Left = 60: .Width = Me.ScaleWidth - (.Left * 2)
        .Top = 60: .Height = Me.ScaleHeight - (.Top * 2) - 300
    End With
    
    With Picture1
        .Left = tab1.ClientLeft
        .Top = tab1.ClientTop
        .Width = tab1.ClientWidth
        .Height = tab1.ClientHeight
    End With
    
    With tReporte
        .Left = tab1.ClientLeft
        .Top = tab1.ClientTop
        .Width = tab1.ClientWidth
        .Height = tab1.ClientHeight
    End With

End Sub

