VERSION 5.00
Begin VB.Form MaDireccion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingreso de Dirección"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Dirección"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3300
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   5295
      Begin VB.ComboBox cCampo2 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2880
         TabIndex        =   17
         Top             =   2040
         Width           =   1455
      End
      Begin VB.ComboBox cCampo1 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ComboBox cComplejo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   960
         Width           =   3495
      End
      Begin VB.ComboBox cLocalidad 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   600
         Width           =   3135
      End
      Begin VB.ComboBox cDepartamento 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox tApartamento 
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   4320
         MaxLength       =   5
         TabIndex        =   14
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox tNumero 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   9
         Top             =   1680
         Width           =   615
      End
      Begin VB.CheckBox cBis 
         Alignment       =   1  'Right Justify
         Caption         =   "B&is:"
         Height          =   285
         Left            =   3120
         TabIndex        =   12
         Top             =   1665
         Width           =   555
      End
      Begin VB.TextBox tLetra 
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   11
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox tSenda 
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   16
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox tBloque 
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   4320
         MaxLength       =   3
         TabIndex        =   18
         Top             =   2040
         Width           =   615
      End
      Begin VB.ComboBox cCalle 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1440
         Style           =   1  'Simple Combo
         TabIndex        =   7
         Top             =   1320
         Width           =   3495
      End
      Begin VB.ComboBox cEntre1 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         Style           =   1  'Simple Combo
         TabIndex        =   20
         Top             =   2520
         Width           =   3495
      End
      Begin VB.ComboBox cEntre2 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         Style           =   1  'Simple Combo
         TabIndex        =   22
         Top             =   2880
         Width           =   3495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Complejo:"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lLocalidad 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Localidad:"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lDepartamento 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Departamento:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lEntre2 
         Alignment       =   1  'Right Justify
         Caption         =   "&Y:"
         Height          =   255
         Left            =   1080
         TabIndex        =   21
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label lEntre1 
         Caption         =   "Ent&re calle:"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lApartamento 
         Alignment       =   1  'Right Justify
         Caption         =   "A&pto:"
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lNumero 
         Alignment       =   1  'Right Justify
         Caption         =   "Nr&o:"
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lCalle 
         Alignment       =   1  'Right Justify
         Caption         =   "Call&e:"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lLetra 
         Caption         =   "Le&tra:"
         Height          =   255
         Left            =   2115
         TabIndex        =   10
         Top             =   1680
         Width           =   495
      End
   End
   Begin VB.TextBox tAmpliacion 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   120
      MaxLength       =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   3720
      Width           =   5295
   End
   Begin VB.TextBox tVive 
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   2400
      MaxLength       =   11
      TabIndex        =   27
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CheckBox cConfirmada 
      Alignment       =   1  'Right Justify
      Caption         =   "Con&firmada"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   4440
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   29
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton bAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   28
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label lAmpiacion 
      BackStyle       =   0  'Transparent
      Caption         =   "A&mpliación:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "&Vive Desde:"
      Height          =   255
      Left            =   1380
      TabIndex        =   26
      Top             =   4440
      Width           =   975
   End
End
Attribute VB_Name = "MaDireccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim RsDir As rdoResultset

Dim iDireccion As Long                   'Codigo de Direccion DirCodigo
Dim iCopia As Long
Dim sTextoDir As String               'Texto Armado de la Direccion

Public Property Get pCodigoDireccion() As Long
    pCodigoDireccion = iDireccion
End Property

Public Property Let pCodigoDireccion(Codigo As Long)
    iDireccion = Codigo
End Property

Public Property Get pCopiaDireccion() As Long
    pCopiaDireccion = iCopia
End Property

Public Property Let pCopiaDireccion(Codigo As Long)
    iCopia = Codigo
End Property

Private Sub AccionGrabar()

    If Not ValidoCampos Then
        Exit Sub
    End If
    
    'PREGUNTO PARA GRABAR----------------------------------------------------------------------------------
    If MsgBox("Confirma almacenar los datos ingresados en la ficha.", vbQuestion + vbYesNo, "GRABAR") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = 11
    
    If iCopia = iDireccion Or iCopia = -1 Then         'Como son las mismas agrego una distinta
        
        On Error GoTo errorBT
        cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
        On Error GoTo errorET
        RsDir.Requery
        
        RsDir.AddNew
        CargoCamposBDDireccion
        RsDir.Update
        
        Cons = "Select Max(DirCodigo) from Direccion"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        iCopia = RsAux(0)
        RsAux.Close
                
        cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------
        RsDir.Requery
        
    Else    'Como son distintas Trabajo con la copia
        On Error GoTo errGrabar
        RsDir.Edit
        CargoCamposBDDireccion
        RsDir.Update
    End If
    
    ArmoDireccion
    
    'Si la calle = N/D registro un suceso Lo grabo fuera de la transaccion --> pero no importa
    On Error Resume Next
    If UCase(cCalle.Text) = "N/D" Then
        RegistroSuceso gFechaServidor, TipoSuceso.Direcciones, paCodigoDeTerminal, InSuceso.pUsuario, 0, _
        Descripcion:="Calle N/D", Defensa:=SacoEnter(sTextoDir)
    End If
        
    Unload Me
    Exit Sub

errGrabar:
    clsError.MuestroError "Ocurrió un error al grabar la información de la dirección."
    Exit Sub

errorBT:
    Screen.MousePointer = 0
    clsError.MuestroError "No se ha podido inicializar la transacción. Reintente la operación."
    RsDir.Requery
    Exit Sub

errorET:
    Resume ErrorRoll

ErrorRoll:
    cBase.RollbackTrans
    RsDir.Requery
    Screen.MousePointer = 0
    clsError.MuestroError "No se ha podido realizar la transacción. Reintente la operación."
End Sub

Private Sub AccionEliminar()
    
On Error GoTo errEliminar

    If iCopia = -1 Then
        'Ya estaba eliminada
        Unload Me
        Exit Sub
    End If
    
    'PREGUNTO PARA ELIMINAR----------------------------------------------------------------------------------
    If MsgBox("Confirma eliminar la dirección de la ficha.", vbQuestion + vbYesNo, "ELIMINAR") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = 11
    
    If iCopia <> iDireccion Then
        'Elimino la Copia
        Cons = "Select * from Direccion Where DirCodigo = " & iCopia
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.Delete
        RsAux.Close
    End If
    
    iCopia = -1     'Mando la copia en -1 para que al grabar elimine la original
    sTextoDir = ""
    Unload Me
    Exit Sub
    
errEliminar:
    clsError.MuestroError "Ocurrió un error al eliminar la información de la dirección."
    Exit Sub
End Sub

Private Sub CargoDatosDireccion()

    'DATOS LOCALIDAD Y DEPARTAMENTO ---------------------------------------------------------------------
    Cons = "Select LocCodigo, LocNombre, LocDepartamento From Calle, Localidad" _
            & " Where CalCodigo = " & RsDir!DirCalle _
            & " And CalLocalidad = LocCodigo"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    'Busco el departamento.------------------------------------------------
    BuscoCodigoEnCombo cDepartamento, RsAux!LocDepartamento
    
    'Cargo la localidad------------------------------------------------------
    Cons = "Select LocCodigo, LocNombre from Localidad Where LocDepartamento = " & RsAux!LocDepartamento
    CargoCombo Cons, cLocalidad, Trim(RsAux!LocNombre)
    
    'Carjo el Complejo Habitacional-----------------------------------------
    If Not IsNull(RsDir!DirComplejo) Then
        Cons = "Select ComCodigo, ComNombre from Complejo Where ComLocalidad = " & RsAux!LocCodigo
        CargoCombo Cons, cComplejo, ""
        BuscoCodigoEnCombo cComplejo, RsDir!DirComplejo
    End If
    
    RsAux.Close '----------------------------------------------------------------------------------------------------
    
    'Inserto la calle en el combo.------------------------------------------
    cCalle.AddItem BuscoCalleCodigo(RsDir!DirCalle)
    cCalle.ItemData(cCalle.NewIndex) = RsDir!DirCalle
    cCalle.ListIndex = 0
    
    'Número de puerta.-----------------------------------------------------
    If Not IsNull(RsDir!DirPuerta) Then tNumero.Text = Trim(RsDir!DirPuerta)
    
    'Número de letra.-----------------------------------------------------
    If Not IsNull(RsDir!DirLetra) Then tLetra.Text = Trim(RsDir!DirLetra)
    
    'Bis.-----------------------------------------------------------------------
    If RsDir!DirBis Then cBis.Value = 1
    
    'Número de apartamento.-------------------------------------------
    If Not IsNull(RsDir!DirApartamento) Then tApartamento.Text = Trim(RsDir!DirApartamento)
    
    'Campo1 (Senda)------------------------------------------------------------------
    If Not IsNull(RsDir!DirCampo1) Then BuscoCodigoEnCombo cCampo1, RsDir!DirCampo1
    If Not IsNull(RsDir!DirSenda) Then tSenda.Text = Trim(RsDir!DirSenda)
    
    'Campo2 (Bloque)----------------------------------------------------------------
    If Not IsNull(RsDir!DirCampo2) Then BuscoCodigoEnCombo cCampo2, RsDir!DirCampo2
    If Not IsNull(RsDir!DirBloque) Then tBloque.Text = Trim(RsDir!DirBloque)
    
    'Entre1.----------------------------------------------------------------
    If Not IsNull(RsDir!DirEntre1) Then
        cEntre1.AddItem BuscoCalleCodigo(RsDir!DirEntre1)
        cEntre1.ItemData(cCalle.NewIndex) = RsDir!DirEntre1
        cEntre1.ListIndex = 0
    End If
    
    'Entre2----------------------------------------------------------------
    If Not IsNull(RsDir!DirEntre2) Then
        cEntre2.AddItem BuscoCalleCodigo(RsDir!DirEntre2)
        cEntre2.ItemData(cCalle.NewIndex) = RsDir!DirEntre2
        cEntre2.ListIndex = 0
    End If
    
    'Ampliacion----------------------------------------------------------------
    If Not IsNull(RsDir!DirAmpliacion) Then tAmpliacion.Text = Trim(RsDir!DirAmpliacion)

    'Confirmada-----------------------------------------------------------------------
    If RsDir!DirConfirmada Then cConfirmada.Value = 1
    
    'Vive desde
    If Not IsNull(RsDir!DirVive) Then tVive.Text = Format(RsDir!DirVive, "d-Mmm-yyyy")
    
End Sub

Private Sub CargoCamposBDDireccion()

    If cComplejo.ListIndex <> -1 Then RsDir!DirComplejo = cComplejo.ItemData(cComplejo.ListIndex)
    RsDir!DirCalle = cCalle.ItemData(cCalle.ListIndex)
    RsDir!DirPuerta = tNumero.Text
    
    If Trim(tLetra.Text) <> "" Then RsDir!DirLetra = Trim(tLetra.Text) Else: RsDir!DirLetra = Null
    If cBis.Value = 0 Then RsDir!DirBis = False Else: RsDir!DirBis = True
    If Trim(tApartamento.Text) <> "" Then RsDir!DirApartamento = Trim(tApartamento.Text) Else: RsDir!DirApartamento = Null
    
    If cCampo1.ListIndex <> -1 Then
        RsDir!DirCampo1 = cCampo1.ItemData(cCampo1.ListIndex)
        If Trim(tSenda.Text) <> "" Then RsDir!DirSenda = Trim(tSenda.Text) Else: RsDir!DirSenda = Null
    Else
        RsDir!DirCampo1 = Null: RsDir!DirSenda = Null
    End If
    
    If cCampo2.ListIndex <> -1 Then
        RsDir!DirCampo2 = cCampo2.ItemData(cCampo2.ListIndex)
        If Trim(tBloque.Text) <> "" Then RsDir!DirBloque = Trim(tBloque.Text) Else: RsDir!DirBloque = Null
    Else
        RsDir!DirCampo2 = Null: RsDir!DirBloque = Null
    End If

    If Trim(cEntre1.Text) <> "" Then RsDir!DirEntre1 = cEntre1.ItemData(cEntre1.ListIndex) Else: RsDir!DirEntre1 = Null
    If Trim(cEntre2.Text) <> "" Then RsDir!DirEntre2 = cEntre2.ItemData(cEntre2.ListIndex) Else: RsDir!DirEntre2 = Null
    
    If Trim(tAmpliacion.Text) <> "" Then RsDir!DirAmpliacion = Trim(tAmpliacion.Text) Else: RsDir!DirAmpliacion = Null
    If cConfirmada.Value = 0 Then RsDir!DirConfirmada = False Else: RsDir!DirConfirmada = True
    If Trim(tVive.Text) <> "" Then RsDir!DirVive = Format(tVive.Text, FormatoFH) Else: RsDir!DirVive = Null
    
End Sub

Private Function ValidoCampos()

    ValidoCampos = False
    'Valido Direccion-----------------------------------------------------------------------------------------------------------
    If cDepartamento.ListIndex = -1 Or cLocalidad.ListIndex = -1 Or cCalle.ListIndex = -1 Or Trim(tNumero.Text) = "" Then
        MsgBox "Los datos ingresados que definen la dirección están incompletos. Verifique", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    
    If Not IsNumeric(tNumero.Text) Then
        MsgBox "Los datos ingresados que definen la dirección no son correctos. Verifique", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    
    If (Trim(cEntre1.Text) <> "" And cEntre1.ListIndex = -1) Or (Trim(cEntre2.Text) <> "" And cEntre2.ListIndex = -1) Then
        MsgBox "Los datos ingresados que definen la dirección no son correctos. Verifique", vbExclamation, "ATENCIÓN"
        Exit Function
    End If

    If Trim(tVive.Text) <> "" Then
        If Not IsDate(tVive.Text) Then
            MsgBox "La fecha ingresada no es correcta. Verifique", vbExclamation, "ATENCIÓN"
            Exit Function
        Else
            If CDate(tVive.Text) > Date Then
                MsgBox "La fecha ingresada no debe ser mayor a la actual", vbExclamation, "ATENCIÓN"
                Exit Function
            End If
        End If
    End If
    
    ValidoCampos = True
    
End Function

Private Function ValidoCamposEliminar()

    ValidoCamposEliminar = False
    'Valido Todo en blanco-----------------------------------------------------------------------------------------------------------
    If Trim(cDepartamento.Text) = "" And Trim(cLocalidad.Text) = "" And Trim(cCalle.Text) = "" And Trim(tNumero.Text) = "" And _
        Trim(tLetra.Text) = "" And Trim(tApartamento.Text) = "" And Trim(tSenda.Text) = "" And Trim(tBloque.Text) = "" And _
        Trim(cEntre1.Text) = "" And Trim(cEntre2.Text) = "" And cBis.Value = 0 And Trim(tAmpliacion.Text) = "" And _
        Trim(tVive.Text) = "" And cConfirmada.Value = 0 Then
        
        ValidoCamposEliminar = True
    End If
    
End Function

Public Property Get pTextoDireccion() As String
    pTextoDireccion = sTextoDir
End Property
Public Property Let pTextoDireccion(Texto As String)
    sTextoDir = Texto
End Property

Private Sub bAceptar_Click()

    'Veo si modifica la direccion o la va a eliminar
    If ValidoCamposEliminar Then
        AccionEliminar
    Else
        AccionGrabar
    End If
    
End Sub

Private Sub cBis_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tApartamento.SetFocus
End Sub

Private Sub cCalle_Change()

    If Trim(tNumero.Text) <> "" Then
        'Borro si la calle que estaba era <> de N/D
        On Error Resume Next
        If cCalle.ListCount > 0 Then If cCalle.List(0) = "N/D" Then Exit Sub
        
        tNumero.Text = ""
        tLetra.Text = ""
        tApartamento.Text = ""
        tSenda.Text = ""
        tBloque.Text = ""
        cBis.Value = 0
        cEntre1.Clear
        cEntre2.Clear
        tAmpliacion.Text = ""
        cConfirmada.Value = 0
        tVive.Text = ""
    End If
    
End Sub

Private Sub cComplejo_Change()
    Selecciono cComplejo, cComplejo.Text, gTecla
End Sub

Private Sub cComplejo_GotFocus()
    cComplejo.SelStart = 0
    cComplejo.SelLength = Len(cComplejo.Text)
End Sub

Private Sub cComplejo_KeyDown(KeyCode As Integer, Shift As Integer)
    gTecla = KeyCode
    gIndice = cComplejo.ListIndex
End Sub

Private Sub cComplejo_KeyPress(KeyAscii As Integer)
    cComplejo.ListIndex = gIndice
    
    If KeyAscii = vbKeyReturn Then
        If cComplejo.ListIndex <> -1 Then
            
            If cCalle.Text = "" Then
                'Cargo la direccion del complejo al cliente-----------------------------------------------
                Screen.MousePointer = 11
                On Error GoTo errDir
                Cons = "Select * from Complejo, Calle" _
                        & " Where ComCodigo = " & cComplejo.ItemData(cComplejo.ListIndex) _
                        & " And ComCalle = CalCodigo"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsAux.EOF Then
                    cCalle.Clear
                    cCalle.AddItem Trim(RsAux!CalNombre)
                    cCalle.ItemData(cCalle.NewIndex) = RsAux!ComCalle
                    cCalle.ListIndex = 0
                    tNumero.Text = RsAux!ComNumero
                End If
                RsAux.Close
                Screen.MousePointer = 0
                '-----------------------------------------------------------------------------------------------
            End If
        End If
        Foco cCalle
    End If
    Exit Sub

errDir:
    Screen.MousePointer = 0
    clsError.MuestroError "Ocurrió un error al buscar la dirección del complejo."
End Sub

Private Sub cComplejo_KeyUp(KeyCode As Integer, Shift As Integer)
    ComboKeyUp cComplejo
End Sub

Private Sub cComplejo_LostFocus()
    cComplejo.SelLength = 0
    gIndice = -1
End Sub

Private Sub cConfirmada_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        tVive.SetFocus
    End If
    
End Sub

Private Sub cDepartamento_Change()

    Selecciono cDepartamento, cDepartamento.Text, gTecla
    cLocalidad.Clear
    LimpioFichaDireccion
    
End Sub

Private Sub cDepartamento_Click()

    cLocalidad.Clear
    LimpioFichaDireccion
    
End Sub

Private Sub cDepartamento_GotFocus()
    
    cDepartamento.SelStart = 0
    cDepartamento.SelLength = Len(cDepartamento.Text)
    
End Sub

Private Sub cDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
        Exit Sub
    End If
    
    gTecla = KeyCode
    gIndice = cDepartamento.ListIndex
    
End Sub

Private Sub cDepartamento_KeyPress(KeyAscii As Integer)

    cDepartamento.ListIndex = gIndice

    If KeyAscii = vbKeyReturn Then
        If cDepartamento.ListIndex <> -1 And cLocalidad.ListIndex = -1 Then
            
            'Cargo las LOCALIDADES
            Cons = "Select LocCodigo, LocNombre From Localidad " _
                    & " Where LocDepartamento = " & cDepartamento.ItemData(cDepartamento.ListIndex) _
                    & " Order by LocNombre"
            CargoCombo Cons, cLocalidad, ""
        End If
        
        cLocalidad.SetFocus
    End If
    
End Sub

Private Sub cDepartamento_KeyUp(KeyCode As Integer, Shift As Integer)

    ComboKeyUp cDepartamento
    
End Sub

Private Sub cDepartamento_LostFocus()

    cDepartamento.SelLength = 0
    gIndice = -1
    
End Sub

Private Sub cEntre1_GotFocus()
    
    cEntre1.SelStart = 0
    cEntre1.SelLength = Len(cEntre1.Text)
    
End Sub

Private Sub cEntre1_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If cLocalidad.ListIndex <> -1 Then
            If ProcesoCalle(Trim(cEntre1.Text), cLocalidad.ItemData(cLocalidad.ListIndex), cEntre1) = 0 Then
                cEntre2.SetFocus
            End If
        Else
            cEntre2.SetFocus
        End If
    End If

End Sub

Private Sub cEntre1_LostFocus()

    cEntre1.SelLength = 0
    
End Sub

Private Sub cEntre2_GotFocus()
    
    cEntre2.SelStart = 0
    cEntre2.SelLength = Len(cEntre2.Text)
    
End Sub

Private Sub cEntre2_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If cLocalidad.ListIndex <> -1 Then
            If ProcesoCalle(Trim(cEntre2.Text), cLocalidad.ItemData(cLocalidad.ListIndex), cEntre2) = 0 Then
                tAmpliacion.SetFocus
            End If
        Else
            tAmpliacion.SetFocus
        End If
    End If

End Sub

Private Sub cEntre2_LostFocus()

    cEntre2.SelLength = 0
    
End Sub

Private Sub cLocalidad_Change()

    Selecciono cLocalidad, cLocalidad.Text, gTecla
    LimpioFichaDireccion
    
End Sub

Private Sub cLocalidad_Click()

    LimpioFichaDireccion
    
End Sub

Private Sub cLocalidad_GotFocus()
    
    cLocalidad.SelStart = 0
    cLocalidad.SelLength = Len(cLocalidad.Text)
    
End Sub

Private Sub cLocalidad_KeyDown(KeyCode As Integer, Shift As Integer)

    gTecla = KeyCode
    gIndice = cLocalidad.ListIndex
    
End Sub

Private Sub cLocalidad_KeyPress(KeyAscii As Integer)

    cLocalidad.ListIndex = gIndice
    
    If KeyAscii = vbKeyReturn Then
    
        If cLocalidad.ListIndex <> -1 And cComplejo.ListIndex = -1 Then
            'Cargo los complejos
            Cons = "Select ComCodigo, Comnombre From Complejo " _
                    & " Where ComLocalidad = " & cLocalidad.ItemData(cLocalidad.ListIndex)
            CargoCombo Cons, cComplejo, ""
        End If
        If cComplejo.ListCount > 0 Then Foco cComplejo Else: Foco cCalle
        
    End If

End Sub

Private Sub cLocalidad_KeyUp(KeyCode As Integer, Shift As Integer)

    ComboKeyUp cLocalidad
    
End Sub

Private Sub cLocalidad_LostFocus()

    cLocalidad.SelLength = 0
    gIndice = -1
    
End Sub

Private Sub cCalle_GotFocus()

    cCalle.SelStart = 0
    cCalle.SelLength = Len(cCalle.Text)
    
End Sub

Private Sub cCalle_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If cLocalidad.ListIndex <> -1 Then
            If ProcesoCalle(Trim(cCalle.Text), cLocalidad.ItemData(cLocalidad.ListIndex), cCalle) = 0 Then
                tNumero.SetFocus
            End If
        Else
            tNumero.SetFocus
        End If
    End If
    
End Sub

Private Sub cCalle_LostFocus()

    cCalle.SelLength = 0
    
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()

On Error GoTo errCargar

    'Cargo los DEPARTAMENTOS
    Cons = "Select DepCodigo, DepNombre From Departamento Order by DepNombre"
    CargoCombo Cons, cDepartamento, ""

    'Cargo los CamposDireccion
    Cons = "Select CDiCodigo, CDiNombre From CamposDireccion"
    CargoCombo Cons, cCampo1, ""
    CargoCombo Cons, cCampo2, ""
    
    Cons = "Select * From Direccion Where DirCodigo = " & iCopia
    Set RsDir = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)

    If iCopia > 0 Then  '-1 Eliminada, 0 No Hay
        CargoDatosDireccion
    
    Else        'Datos por defecto
        'Busco el departamento.------------------------------------------------
        BuscoCodigoEnCombo cDepartamento, paDepartamento
        If cDepartamento.ListIndex <> -1 Then
            'Cargo la localidad------------------------------------------------------
            Cons = "Select LocCodigo, LocNombre from Localidad Where LocDepartamento = " & paDepartamento
            CargoCombo Cons, cLocalidad, ""
            BuscoCodigoEnCombo cLocalidad, paLocalidad
        End If
    End If
    Exit Sub

errCargar:
    clsError.MuestroError "Ocurrió un error al cargar los datos de dirección."
End Sub

Private Sub Form_Unload(Cancel As Integer)

    RsDir.Close
    
End Sub

Private Sub ArmoDireccion()

Dim aTexto As String
    
    aTexto = Trim(cDepartamento.Text) & ", " & Trim(cLocalidad.Text) & Chr(vbKeyReturn) & Chr(10)
    If cComplejo.ListIndex <> -1 Then aTexto = aTexto & Trim(cComplejo.Text) & Chr(vbKeyReturn) & Chr(10)
    aTexto = aTexto & Trim(cCalle.Text) & " " & Trim(tNumero.Text)
    
    If Trim(tLetra.Text) <> "" Then aTexto = aTexto & Trim(tLetra.Text)
    If Trim(tApartamento.Text) <> "" Then aTexto = aTexto & "/" & Trim(tApartamento.Text)
    If cBis.Value = 1 Then aTexto = aTexto & " Bis"
    
    If cCampo1.ListIndex <> -1 Then
        aTexto = aTexto & " " & Trim(cCampo1.Text)
        If Trim(tSenda.Text) <> "" Then aTexto = aTexto & " " & Trim(tSenda.Text)
    End If
    If cCampo2.ListIndex <> -1 Then
        aTexto = aTexto & " " & Trim(cCampo2.Text)
        If Trim(tBloque.Text) <> "" Then aTexto = aTexto & " " & Trim(tBloque.Text)
    End If
    
    If Trim(cEntre1.Text) <> "" Or Trim(cEntre2.Text) <> "" Then
        aTexto = aTexto & Chr(vbKeyReturn) & Chr(10)
        If Trim(cEntre1.Text) <> "" And Trim(cEntre2.Text) <> "" Then
            aTexto = aTexto & "Entre " & Trim(cEntre1.Text) & " y " & Trim(cEntre2.Text)
        Else
            If Trim(cEntre1.Text) <> "" Then
                aTexto = aTexto & "Esquina " & Trim(cEntre1.Text)
            Else
                aTexto = aTexto & "Esquina " & Trim(cEntre2.Text)
            End If
        End If
    End If
    
    If Trim(tAmpliacion.Text) <> "" Then aTexto = aTexto & Chr(vbKeyReturn) & Chr(10) & Trim(tAmpliacion.Text)
    
    If cConfirmada.Value = 1 Then
        aTexto = aTexto & Chr(vbKeyReturn) & Chr(10) & Chr(vbKeyReturn) & Chr(10) & "(Confirmada"
    Else
        aTexto = aTexto & Chr(vbKeyReturn) & Chr(10) & Chr(vbKeyReturn) & Chr(10) & "(No Confirmada"
    End If
    
    If Trim(tVive.Text) <> "" Then aTexto = aTexto & ", Vive desde " & Format(tVive.Text, "Mmm-yyyy")
    
    sTextoDir = aTexto & ")"
    
End Sub

Private Sub Label1_Click()
    Foco tVive
End Sub

Private Sub lAmpiacion_Click()
    Foco tAmpliacion
End Sub

Private Sub lApartamento_Click()
    Foco tApartamento
End Sub

Private Sub lCalle_Click()
    Foco cCalle
End Sub

Private Sub lDepartamento_Click()
    Foco cDepartamento
End Sub

Private Sub lEntre1_Click()
    Foco cEntre1
End Sub

Private Sub lEntre2_Click()
    Foco cEntre2
End Sub

Private Sub lLetra_Click()
    Foco tLetra
End Sub

Private Sub lLocalidad_Click()
    Foco cLocalidad
End Sub

Private Sub lNumero_Click()
    Foco tNumero
End Sub

Private Sub tAmpliacion_GotFocus()
    tAmpliacion.SelStart = 0
    tAmpliacion.SelLength = Len(tAmpliacion.Text)
End Sub

Private Sub tAmpliacion_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = Empty
        cConfirmada.SetFocus
    End If

End Sub

Private Sub tApartamento_GotFocus()

    tApartamento.SelStart = 0
    tApartamento.SelLength = Len(tApartamento.Text)
    
End Sub

Private Sub tApartamento_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Foco cCampo1
    
End Sub

Private Sub tBloque_GotFocus()
    
    tBloque.SelStart = 0
    tBloque.SelLength = Len(tBloque.Text)
    
End Sub

Private Sub tBloque_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        cEntre1.SetFocus
    End If
    
End Sub

Private Sub tLetra_GotFocus()

    tLetra.SelStart = 0
    tLetra.SelLength = Len(tLetra.Text)
    
End Sub

Private Sub tLetra_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then
        cBis.SetFocus
    End If
    
End Sub

Private Sub tNumero_GotFocus()

    tNumero.SelStart = 0
    tNumero.SelLength = Len(tNumero.Text)
    
End Sub

Private Sub tNumero_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        tLetra.SetFocus
    End If

End Sub

Private Sub LimpioFichaDireccion()

    cComplejo.Clear
    cCalle.Clear
    tNumero.Text = ""
    tLetra.Text = ""
    tApartamento.Text = ""
    cCampo1.Text = ""
    tSenda.Text = ""
    cCampo2.Text = ""
    tBloque.Text = ""
    cBis.Value = 0
    cEntre1.Clear
    cEntre2.Clear
    tAmpliacion.Text = ""
    cConfirmada.Value = 0
    tVive.Text = ""
    
End Sub

Private Sub tSenda_GotFocus()

    tSenda.SelStart = 0
    tSenda.SelLength = Len(tSenda.Text)
    
End Sub

Private Sub tSenda_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Foco cCampo2

End Sub

Private Sub tVive_GotFocus()

    tVive.SelStart = 0
    tVive.SelLength = Len(tVive.Text)
    
End Sub

Private Sub tVive_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        bAceptar.SetFocus
    End If

End Sub

Private Sub tVive_LostFocus()

    On Error GoTo errFecha
    If UCase(Left(tVive.Text, 1)) = "A" Then
        If IsNumeric(Mid(tVive.Text, 2, Len(tVive.Text))) Then
            tVive.Text = Format(Date, "d-Mmm-") & (Year(Date) - Mid(tVive.Text, 2, Len(tVive.Text)))
        End If
    End If
    
    If IsDate(tVive.Text) Then tVive.Text = Format(tVive.Text, "d-Mmm-yyyy")
    Exit Sub

errFecha:
End Sub

Private Sub cCampo1_Change()
    Selecciono cCampo1, cCampo1.Text, gTecla
End Sub

Private Sub cCampo1_GotFocus()
    cCampo1.SelStart = 0
    cCampo1.SelLength = Len(cCampo1.Text)
End Sub

Private Sub cCampo1_KeyDown(KeyCode As Integer, Shift As Integer)
    gTecla = KeyCode
    gIndice = cCampo1.ListIndex
End Sub

Private Sub cCampo1_KeyPress(KeyAscii As Integer)
    cCampo1.ListIndex = gIndice
    If KeyAscii = vbKeyReturn Then Foco tSenda
End Sub

Private Sub cCampo1_KeyUp(KeyCode As Integer, Shift As Integer)
    ComboKeyUp cCampo1
End Sub

Private Sub cCampo1_LostFocus()
    cCampo1.SelLength = 0
    gIndice = -1
End Sub

Private Sub cCampo2_Change()
    Selecciono cCampo2, cCampo2.Text, gTecla
End Sub

Private Sub cCampo2_GotFocus()
    cCampo2.SelStart = 0
    cCampo2.SelLength = Len(cCampo2.Text)
End Sub

Private Sub cCampo2_KeyDown(KeyCode As Integer, Shift As Integer)
    gTecla = KeyCode
    gIndice = cCampo2.ListIndex
End Sub

Private Sub cCampo2_KeyPress(KeyAscii As Integer)
    cCampo2.ListIndex = gIndice
    If KeyAscii = vbKeyReturn Then Foco tBloque
End Sub

Private Sub cCampo2_KeyUp(KeyCode As Integer, Shift As Integer)
    ComboKeyUp cCampo2
End Sub

Private Sub cCampo2_LostFocus()
    cCampo2.SelLength = 0
    gIndice = -1
End Sub

