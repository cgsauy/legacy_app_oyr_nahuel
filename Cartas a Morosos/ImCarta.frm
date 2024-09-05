VERSION 5.00
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmImCarta 
   Caption         =   "Impresión de Cartas"
   ClientHeight    =   7605
   ClientLeft      =   1620
   ClientTop       =   2475
   ClientWidth     =   10575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ImCarta.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   10575
   Begin VB.CheckBox chPrevia 
      Caption         =   "Ver Impresión Previa"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   3255
   End
   Begin RichTextLib.RichTextBox cRT 
      Height          =   7335
      Left            =   3720
      TabIndex        =   13
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   12938
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"ImCarta.frx":0442
   End
   Begin VB.CommandButton bImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton bRefresh 
      Caption         =   "Cargar Datos"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ComboBox cPrinter 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selección de Impresiones"
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   1500
      Width           =   3375
      Begin VB.TextBox tCantidad 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   8
         Top             =   540
         Width           =   615
      End
      Begin VB.OptionButton oParcial 
         Caption         =   "&Parcial"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   580
         Width           =   855
      End
      Begin VB.OptionButton oTotal 
         Caption         =   "&Total"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad de Cartas:"
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
   End
   Begin vsViewLib.vsPrinter vsP 
      Height          =   6135
      Left            =   3900
      TabIndex        =   9
      Top             =   360
      Width           =   5775
      _Version        =   196608
      _ExtentX        =   10186
      _ExtentY        =   10821
      _StockProps     =   229
      Appearance      =   1
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ConvInfo        =   1418783674
      PageBorder      =   0
      PhysicalPage    =   -1  'True
      Zoom            =   100
   End
   Begin VB.Label lStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   3660
      Width           =   3375
   End
   Begin VB.Label Label88 
      BackStyle       =   0  'Transparent
      Caption         =   "&Impresora:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total de Cartas a Imprimir:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1140
      Width           =   2055
   End
   Begin VB.Label lTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "50"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   1140
      Width           =   495
   End
End
Attribute VB_Name = "frmImCarta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type typTmp
    Archivo As String
    Data As String
End Type

Dim arrTemplates() As typTmp

Dim aTexto As String

Dim RsDir As rdoResultset
Dim RsAux2 As rdoResultset
Dim rsVal As rdoResultset

Private Sub bImprimir_Click()
    AccionImprimir
End Sub

Private Sub bRefresh_Click()
    AccionRefrescar
End Sub

Private Sub cPrinter_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then oTotal.SetFocus
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()

    On Error GoTo ErrLoad
    lTotal.Caption = "S/D"
    tCantidad.Text = ""
    tCantidad.Enabled = False
    tCantidad.BackColor = Inactivo
    
    'Cargo las impresoras definidas en el sistema-------------------
    Dim X As Printer
    For Each X In Printers
        cPrinter.AddItem Trim(X.DeviceName)
    Next
    If cPrinter.ListCount > 0 Then cPrinter.ListIndex = 0
    '----------------------------------------------------------------------
    
    ReDim arrTemplates(0)
    AccionRefrescar
    
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Error al inicializar el formulario.", Err.Description
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    With vsP
        .Width = Me.ScaleWidth - .Left - 100
        .Height = Me.ScaleHeight - .Top - 100
    End With
    
    With cRT
        .Width = Me.ScaleWidth - .Left - 100
        .Height = Me.ScaleHeight - .Top - 100
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EndMain
End Sub

Private Sub AccionRefrescar()
    
    On Error GoTo errError
    FechaDelServidor
    
    cons = "Select Count(*) from CartaMoroso Where CMoImpreso Is Null"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux(0)) Then lTotal.Caption = rsAux(0) Else: lTotal.Caption = 0
    End If
    rsAux.Close
    
    If Val(lTotal.Caption) = 0 Then MsgBox "No hay cartas almacenadas para imprimir.", vbExclamation, "ATENCIÓN"
    
    Exit Sub
    
errError:
    clsGeneral.OcurrioError "Error al procesar la información.", Err.Description
    Screen.MousePointer = 0
End Sub

Sub AccionImprimir()

    If Val(lTotal.Caption) = 0 Then
        MsgBox "No hay cartas almacenadas para imprimir.", vbExclamation, "No Hay Cartas"
        Exit Sub
    End If
    
    If oParcial.Value Then
        If Not IsNumeric(tCantidad.Text) Or Val(tCantidad.Text) >= Val(lTotal.Caption) Then
            MsgBox "La cantidad de cartas ingresadas no es correcta.", vbExclamation, "ATECIÓN"
            Foco tCantidad: Exit Sub
        End If
    End If
    
    If cPrinter.ListIndex = -1 Then
        MsgBox "Seleccione la impresora para listar las cartas.", vbExclamation, "ATENCIÓN"
        cPrinter.SetFocus: Exit Sub
    End If
    
    If chPrevia.Value = vbUnchecked Then
        MsgBox "Valor 'Ver Impresión Previa' no seleccionado !!. " & vbCrLf & _
                    "Las cartas se van a imprimir y actualizar con la fecha de impresión."
    End If
    
    If MsgBox("Confirma realizar la impresión de las cartas.", vbQuestion + vbYesNo, "Imprimir Cartas") = vbNo Then Exit Sub
    
    vsP.Device = cPrinter.Text
    
    Dim X As Printer
    For Each X In Printers
        If cPrinter.Text = Trim(X.DeviceName) Then
            Set Printer = X
            Exit For
        End If
    Next
    '----------------------------------------------------------------------
    
    CargoDatosAWord

End Sub

Private Sub CargoDatosAWord()

    On Error GoTo Error
    Dim mUltimoPago As String, mNoImpresos As String, bVinoAPagar As Boolean
    Dim mFileText As String
    
    Screen.MousePointer = 11
    
    Dim RsAC As rdoResultset

    Dim aMaximo As Long
    Dim aCMonedaA As Long, aTMonedaA As String
    Dim m_NombreS As String, m_NombreC As String, m_CiRuc As String
    Dim m_IDDireccion As Long
    Dim m_Calle As String, m_Entre As String, m_Localidad As String
    
    aMaximo = 0
    aCMonedaA = 0
    mNoImpresos = ""
    lStatus.Caption = "Imprimiendo ...": lStatus.Refresh
    
    Dim aVoy As Long, aQTotal As Long
    cons = " Select * from CartaMoroso, Cartas  " _
            & " Where CMoImpreso Is Null" _
            & " And CMoCarta = CarCodigo"
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    I = 0
    aVoy = 0
    If oParcial.Value Then aQTotal = CLng(tCantidad.Text) Else aQTotal = CLng(lTotal.Caption)
    
    Do While Not rsAux.EOF
        aVoy = aVoy + 1
        lStatus.Caption = "Imprimiendo ... " & aVoy & " de " & aQTotal
        lStatus.Refresh
        
        
        'Creo el Template al archivo de la carta -> .dot
        mFileText = arr_Template(CStr(Trim(rsAux!CarArchivo)))
        
        'CARGO LOS CAMPOS DEL ARCHIVO-------------------------------------------------------------------!!!!!!!
        
        'Nombres del Cliente-------------------------------------------------------------------------------------------
        cons = "Select CliCiRuc, CliDireccion Direccion, CPeSexo Sexo, " _
                         & " NombreC = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2), " _
                         & " NombreS = (RTrim(CPeNombre1) + ' ' + RTrim(CPeApellido1)) " _
                & " From Cliente, CPersona " _
                & " Where CliCodigo = " & rsAux!CMoCliente _
                & " And CliCodigo = CPeCliente" _
                             & " Union All " _
                & " Select CliCiRuc, CliDireccion Direccion, Sexo = Null, " _
                          & " NombreC = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')'), " _
                          & " NombreS = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')')" _
                & " From Cliente, CEmpresa " _
                & " Where CliCodigo = " & rsAux!CMoCliente _
                & " And CliCodigo = CEmCliente"
                
        Set RsAC = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
        m_NombreS = Trim(RsAC!NombreS)
        m_NombreC = Trim(RsAC!NombreC)
        m_CiRuc = ""
        
        If Not IsNull(RsAC!Sexo) Then
            If RsAC!Sexo = "F" Then m_NombreS = "Sra. " & m_NombreS Else m_NombreS = "Sr. " & m_NombreS
            If RsAC!Sexo = "F" Then m_NombreC = "Sra. " & m_NombreC Else m_NombreC = "Sr. " & m_NombreC
            
            If Not IsNull(RsAC!CliCiRuc) Then m_CiRuc = clsGeneral.RetornoFormatoCedula(Trim(RsAC!CliCiRuc))
        Else
            If Not IsNull(RsAC!CliCiRuc) Then m_CiRuc = clsGeneral.RetornoFormatoRuc(Trim(RsAC!CliCiRuc))
        End If
        
        If Not IsNull(RsAC!Direccion) Then m_IDDireccion = RsAC!Direccion Else m_IDDireccion = 0
        RsAC.Close
        
        'Valido Si vino a pagar despues de almacenada la carta  -----------------------------------------------
        bVinoAPagar = False
        mUltimoPago = Format(rsAux!CMoFAtraso, "yyyy/mm/dd")
        Dim mClienteValidar As Long
        mClienteValidar = rsAux!CMoCliente
        If Not IsNull(rsAux!CMoGarantiaDe) Then mClienteValidar = rsAux!CMoGarantiaDe
        
        If Not fnc_ValidoUltimoPago(mClienteValidar, mUltimoPago) Then
            bVinoAPagar = True
            mNoImpresos = mNoImpresos & _
                                    mUltimoPago & "    " & m_CiRuc & "  " & m_NombreC & vbCrLf
            GoTo etqNext
        End If
        '------------------------------------------------------------------------------------------------------------------
        
        'Cambio las formulas en el template -----------------------------------------------------------------------
        mFileText = Replace(mFileText, "[CiRuc]", m_CiRuc)
        mFileText = Replace(mFileText, "[NombreCompleto]", m_NombreC)
        mFileText = Replace(mFileText, "[NombreSimple]", m_NombreS)
        
        mFileText = Replace(mFileText, "[Fecha]", CStr(Format(Date, "Long Date")))
               
        'Dirección----------------------------------------------------------
        BuscoDatosDireccion m_IDDireccion, m_Calle, m_Entre, m_Localidad
        
        mFileText = Replace(mFileText, "[Direccion]", Trim(m_Calle))
        mFileText = Replace(mFileText, "[EntreCalle]", Trim(m_Entre))
        mFileText = Replace(mFileText, "[Localidad]", Trim(m_Localidad))
        
        'Fecha de Atraso----------------------------------------------------
        mFileText = Replace(mFileText, "[FechaAtraso]", CStr(Format(rsAux!CMoFAtraso, "d Mmmm, yyyy")))
        
        'Deuda----------------------------------------------------------------------------------------------------------
        If aCMonedaA <> rsAux!CMoMoneda Then
            aTMonedaA = BuscoSignoMoneda(rsAux!CMoMoneda)
            aCMonedaA = rsAux!CMoMoneda
        End If
        mFileText = Replace(mFileText, "[Deuda]", Format(rsAux!CMoDeuda, FormatoMonedaP))
                
        'Garantia---------------------------------------------------------------------------------------------------------
        m_NombreC = ""
        If Not IsNull(rsAux!CMoGarantiaDe) Then
            cons = " Select Tipo = " & 1 & ", CPeSexo Sexo, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2) " _
                    & " From CPersona Where CPeCliente = " & rsAux!CMoGarantiaDe _
                    & " UNION ALL " _
                    & " Select Tipo = " & 2 & ", CEmFantasia Sexo, Nombre = CEmNombre " _
                    & " From CEmpresa Where CEmCliente = " & rsAux!CMoGarantiaDe
            Set RsAC = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
            
            If RsAC!Tipo = 1 Then m_NombreC = Trim(RsAC!Nombre)
            If RsAC!Tipo = 2 Then
                If Not IsNull(RsAC!Nombre) Then m_NombreC = Trim(RsAC!Nombre) Else m_NombreC = Trim(RsAC!Sexo)
            End If
            RsAC.Close
        End If
        
        mFileText = Replace(mFileText, "[GarantiaDe]", m_NombreC)
        '---------------------------------------------------------------------------------------------------------------
                        
        cRT.TextRTF = mFileText
        
        If chPrevia.Value = vbUnchecked Then cRT.SelPrint Printer.hDC
        
etqNext:
        If chPrevia.Value = vbUnchecked Then
            If Not bVinoAPagar Then
                'ACTUALIZO CON FECHA DE IMPRESION ---------------------------------------------------------------
                cons = "Update CartaMoroso Set CMoImpreso = '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "'" _
                       & " Where CMoCodigo = " & rsAux!CMoCodigo
                cBase.Execute cons
                '----------------------------------------------------------------------------------------------------------
            Else
                'ELIMINO REG de LA CARTA                 ---------------------------------------------------------------
                cons = "Delete CartaMoroso Where CMoCodigo = " & rsAux!CMoCodigo
                cBase.Execute cons
                '----------------------------------------------------------------------------------------------------------
            End If
        End If
        
        If oParcial.Value Then
            I = I + 1
            If I >= Val(tCantidad.Text) Then Exit Do
        End If
        
        rsAux.MoveNext
    Loop
    
    rsAux.Close
    
    lTotal.Caption = "S/D"
    tCantidad.Text = ""
    
    If Trim(mNoImpresos) <> "" Then
        mNoImpresos = "CARTAS NO IMPRESAS" & vbCrLf & vbCrLf & mNoImpresos
        cRT.Text = mNoImpresos
        If chPrevia.Value = vbUnchecked Then cRT.SelPrint Printer.hDC
    End If
    
    Screen.MousePointer = 0
    
    lStatus.Caption = "Impresión finalizada."
    Exit Sub
    
Error:
    clsGeneral.OcurrioError "Error al realizar la impresión de las cartas.", CStr(Trim(Err.Description))
    lTotal.Caption = "S/D"
    tCantidad.Text = ""
    Screen.MousePointer = 0
End Sub

'----------------------------------------------------------------------------------------------------------------------------
'   Busca la direccion del cliente con codigo: Codigo
'   Valores de Retorno: rCalle, rEntre, rLocalidad
'----------------------------------------------------------------------------------------------------------------------------
Private Sub BuscoDatosDireccion(Codigo As Long, rCalle As String, rEntre As String, rLocalidad As String)
    
    cons = "Select Direccion.*, LocNombre, DepNombre, CalNombre From Direccion, Calle, Localidad, Departamento" _
            & " Where DirCodigo = " & Codigo _
            & " And DirCalle = CalCodigo And CalLocalidad = LocCodigo" _
            & " And LocDepartamento = DepCodigo"
    
    Set RsDir = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If RsDir.EOF Then
        rLocalidad = "N/D": rEntre = "N/D": rCalle = "N/D"
        RsDir.Close: Exit Sub
    End If
    
    rLocalidad = Trim(RsDir!DepNombre) & ", " & Trim(RsDir!LocNombre)
        
    'Proceso la Calle--------------------------------------------------------------------------------------------------------------
    aTexto = Trim(RsDir!CalNombre) & " "
    If Trim(RsDir!DirPuerta) = 0 Then aTexto = aTexto & "S/N" Else: aTexto = aTexto & Trim(RsDir!DirPuerta)
    If Not IsNull(RsDir!DirLetra) Then aTexto = aTexto & Trim(RsDir!DirLetra)
    If Not IsNull(RsDir!DirApartamento) Then aTexto = aTexto & "/" & Trim(RsDir!DirApartamento)
    If RsDir!DirBis Then aTexto = aTexto & " Bis"
    
    
    'Campo 1 de la Direccion---------------------------------------------------------------------------------------
    If Not IsNull(RsDir!DirCampo1) Then
        cons = "Select CDiAbreviacion from CamposDireccion Where CDiCodigo = " & RsDir!DirCampo1
        Set RsAux2 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux2.EOF Then
            aTexto = aTexto & " " & Trim(RsAux2!CDiAbreviacion)
            If Not IsNull(RsDir!DirSenda) Then aTexto = aTexto & " " & Trim(RsDir!DirSenda)
        End If
        RsAux2.Close
    End If
    'Campo 2 de la Direccion---------------------------------------------------------------------------------------
    If Not IsNull(RsDir!DirCampo2) Then
        cons = "Select CDiAbreviacion from CamposDireccion Where CDiCodigo = " & RsDir!DirCampo2
        Set RsAux2 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux2.EOF Then
            aTexto = aTexto & " " & Trim(RsAux2!CDiAbreviacion)
            If Not IsNull(RsDir!DirBloque) Then aTexto = aTexto & " " & Trim(RsDir!DirBloque)
        End If
        RsAux2.Close
    End If
    '---------------------------------------------------------------------------------------------------------------------
        
    'Saco CP de la ZONA------------------------------------------------------------------------------------------
    cons = "Select ZonCPostal from CalleZona, Zona" _
           & " Where CZoCalle = " & RsDir!DirCalle _
           & " And CZoDesde <= " & RsDir!DirPuerta _
           & " And CZoHasta >= " & RsDir!DirPuerta _
           & " And CZoZona = ZonCodigo"
    Set RsAux2 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux2.EOF Then
        If Not IsNull(RsAux2!ZonCPostal) Then aTexto = aTexto & "  CP " & Trim(RsAux2!ZonCPostal)
    End If
    RsAux2.Close
    '------------------------------------------------------------------------------------------------------------
    rCalle = Trim(aTexto)
    
    'Proceso la Entre Calles---------------------------------------------------------------------------------------------
    rEntre = ""
    If Not IsNull(RsDir!DirEntre1) Or Not IsNull(RsDir!DirEntre2) Then
        If Not IsNull(RsDir!DirEntre1) And Not IsNull(RsDir!DirEntre2) Then
            
            cons = "Select * from Calle where CalCodigo = " & RsDir!DirEntre1
            Set RsAux2 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux2.EOF Then aTexto = "Entre " & Trim(RsAux2!CalNombre)
            RsAux2.Close
            
            cons = "Select * from Calle where CalCodigo = " & RsDir!DirEntre2
            Set RsAux2 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux2.EOF Then aTexto = aTexto & " y " & Trim(RsAux2!CalNombre)
            RsAux2.Close
            
        Else
            If Not IsNull(RsDir!DirEntre1) Then
                cons = "Select * from Calle where CalCodigo = " & RsDir!DirEntre1
                Set RsAux2 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
                If Not RsAux2.EOF Then aTexto = "Esquina " & Trim(RsAux2!CalNombre)
                RsAux2.Close
            Else
                cons = "Select * from Calle where CalCodigo = " & RsDir!DirEntre2
                Set RsAux2 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
                If Not RsAux2.EOF Then aTexto = "Esquina " & Trim(RsAux2!CalNombre)
                RsAux2.Close
            End If
        End If
        rEntre = aTexto
    End If
    
    RsDir.Close

End Sub

Private Sub oParcial_Click()
    tCantidad.Enabled = True
    tCantidad.BackColor = vbWindowBackground
End Sub

Private Sub oTotal_Click()
    tCantidad.Enabled = False
    tCantidad.BackColor = Inactivo
End Sub

Private Function BuscoSignoMoneda(Codigo As Variant) As String
On Error GoTo ErrBU
    
Dim Rs As rdoResultset

    BuscoSignoMoneda = ""

    cons = "SELECT * FROM Moneda WHERE MonCodigo = " & Codigo
    Set Rs = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not Rs.EOF Then BuscoSignoMoneda = Trim(Rs!MonSigno)
    Rs.Close
    Exit Function
    
ErrBU:
End Function


Private Function fnc_Load(strFile As String) As String

    On Error GoTo PROC_ERR
    fnc_Load = ""
    cRT.LoadFile strFile
    fnc_Load = cRT.TextRTF
    Exit Function
  
PROC_ERR:
  MsgBox "Error al cargar matriz." & vbCrLf & _
               Err.Number & "- " & Err.Description, vbCritical, "Archivo: " & strFile
End Function


Private Function fnc_Load2(strFile As String) As String

Dim intFile As Integer, varTmp As String
Dim mFile As String

    On Error GoTo PROC_ERR
    fnc_Load2 = ""
    intFile = FreeFile
    Open strFile For Input As intFile
    
    Do Until EOF(intFile)
        Input #intFile, varTmp
        mFile = mFile & varTmp
    Loop
    
    Close #intFile
     
    fnc_Load2 = mFile
    Exit Function
  
PROC_ERR:
  MsgBox "Error al cargar matriz." & vbCrLf & _
               Err.Number & "- " & Err.Description, vbCritical, "Archivo: " & strFile
  Close #intFile
End Function

Private Function arr_Template(mNameDoc As String) As String

    'Lo Trucho Temporal ------------------------------------------------------------
    If UCase(Right(mNameDoc, 3)) = "DOT" Then
        mNameDoc = Replace(UCase(mNameDoc), ".DOT", ".RTF")
    End If
    
    arr_Template = ""
    Dim I As Integer, idArr As Integer, bOk As Boolean
    bOk = False
    
    For I = LBound(arrTemplates) To UBound(arrTemplates)
        With arrTemplates(I)
            
            If I = 0 And Trim(.Archivo) = "" Then
                .Archivo = Trim(mNameDoc)
                .Data = fnc_Load(mNameDoc)
                arr_Template = .Data
                bOk = True
                Exit For
            End If
            
            If Trim(.Archivo) = mNameDoc Then
                arr_Template = .Data
                bOk = True
                Exit For
            End If
        
        End With
    Next
    
    If bOk Then Exit Function
    
    I = UBound(arrTemplates) + 1
    ReDim Preserve arrTemplates(I)
    
    With arrTemplates(I)
        .Archivo = Trim(mNameDoc)
        .Data = fnc_Load(mNameDoc)
        arr_Template = .Data
    End With

End Function

Private Function fnc_ValidoUltimoPago(mIDCliente As Long, UP_ymd As String) As Boolean
On Error GoTo errVUP

    fnc_ValidoUltimoPago = False
    
    cons = "Select Min(CreProximoVto) From  Credito, Documento" & _
                " Where CreFactura = DocCodigo" & _
                " And CreSaldoFactura > 0" & _
                " And CreTipo = 0 And DocAnulado = 0" & _
                " And CreCliente = " & mIDCliente
    
    Set rsVal = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsVal.EOF Then
        If Not IsNull(rsVal(0)) Then
            If Format(rsVal(0), "yyyy/mm/dd") = UP_ymd Then
                fnc_ValidoUltimoPago = True
            Else
                UP_ymd = Format(rsVal(0), "dd/mm/yyyy")
            End If
        Else
            UP_ymd = "--------"
        End If
    Else
        UP_ymd = "--------"
    End If
    rsVal.Close
            
errVUP:
End Function
