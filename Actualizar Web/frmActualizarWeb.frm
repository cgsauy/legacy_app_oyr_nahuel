VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmActualizarWeb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizar Web"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmActualizarWeb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmOpcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Caption         =   "Opciones"
      ForeColor       =   &H00800000&
      Height          =   2295
      Left            =   120
      TabIndex        =   23
      Top             =   840
      Width           =   5535
      Begin VB.CheckBox chArchivo 
         Appearance      =   0  'Flat
         Caption         =   "Archivo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   24
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chAccesorios 
         Appearance      =   0  'Flat
         Caption         =   "Accesorios"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   14
         Top             =   1920
         Width           =   1455
      End
      Begin VB.OptionButton opHacer 
         Appearance      =   0  'Flat
         Caption         =   "Solo un artículo"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   2640
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox tArticulo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Top             =   720
         Width           =   4455
      End
      Begin VB.CheckBox chLista 
         Appearance      =   0  'Flat
         Caption         =   "Listas de Precios generales"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   2535
      End
      Begin VB.CheckBox chPrecio 
         Appearance      =   0  'Flat
         Caption         =   "Precios"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   12
         Top             =   1560
         Width           =   855
      End
      Begin VB.CheckBox chGlosario 
         Appearance      =   0  'Flat
         Caption         =   "&Glosario"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CheckBox chPlanes 
         Appearance      =   0  'Flat
         Caption         =   "&Cuotas y Planes"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox chGrupo 
         Appearance      =   0  'Flat
         Caption         =   "&Grupos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox chMarca 
         Appearance      =   0  'Flat
         Caption         =   "&Marca"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox chTipo 
         Appearance      =   0  'Flat
         Caption         =   "&Tipo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton opHacer 
         Appearance      =   0  'Flat
         Caption         =   "Parcial"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton opHacer 
         Appearance      =   0  'Flat
         Caption         =   "Total"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox chEspecie 
         Appearance      =   0  'Flat
         Caption         =   "&Especies"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox chStock 
         Appearance      =   0  'Flat
         Caption         =   "Stock"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   8
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "A&rtículo"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.TextBox tSuceso 
      Appearance      =   0  'Flat
      Height          =   1935
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Text            =   "frmActualizarWeb.frx":030A
      Top             =   960
      Width           =   5535
   End
   Begin VB.CommandButton bSucesos 
      Caption         =   "S&ucesos"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar pbParcial 
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton bSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton bActualizar 
      Caption         =   "&Actualizar"
      Height          =   375
      Left            =   3360
      TabIndex        =   16
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lLogoDet 
      BackStyle       =   0  'Transparent
      Caption         =   "Paso 1 de 10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   21
      Top             =   480
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmActualizarWeb.frx":0310
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lLogoPaso 
      BackStyle       =   0  'Transparent
      Caption         =   "Paso 1 de 10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   20
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00CC9966&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   735
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frmActualizarWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Modificaciones
'6/6/203 Volvi a poner la situación de ArtHabilitado antes tenía siempre true.
'22/12/03 agregue cambios x combos (un art. que es combo tiene que tener precio)

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim rsAr As rdoResultset
Private Type tPlantilla
    idArticulo As Long
    idPlaWeb As Long
    idPlaIntra As Long
End Type

Private Type tArtPrecio
    idArt As Long
    Ctdo As Currency
    CuotaFinanciado As Currency
    TCuotaAbrev As String
    Plan As String
End Type

Private Type tCodigoNombre
    Codigo As Long
    Nombre As String
End Type

Private Type tArticulo
    ID As Long
    Sanos As Long
    Vta As Long
End Type


Private Type tArtHtmAsp
    sPath As String
    sData As String
End Type

Private arrFileWeb() As tArtHtmAsp
Private arrFileIntra() As tArtHtmAsp

Private colError As Collection
Private sArtNew As String

'Private arrTipoPlan() As tCodigoNombre
Private arrArtPla() As tPlantilla
Private arrArtPre() As tArtPrecio

Private lCantArtMod As Long
Private boCancel As Boolean      'Señal que me indica si el usuario desea cancelar.
Private bload As Boolean
Private sArtMod As String
Private sTipoEnCGSA As String

Private Function f_SetFilterFind(ByVal sTexto As String) As String
    sTexto = RTrim(sTexto)
    sTexto = Replace(sTexto, " ", "%")
    sTexto = Replace(sTexto, "*", "%")
    f_SetFilterFind = "'" & sTexto & "%'"
End Function

Private Sub bActualizar_Click()
    
    If bActualizar.Caption = "&Actualizar" Then
    
        If Not bCentinela Then
            If opHacer(2).Value And Val(tArticulo.Tag) = 0 Then
                MsgBox "Seleccione un artículo.", vbExclamation, "Validación"
                tArticulo.SetFocus
                Exit Sub
            End If
            
            If MsgBox("¿Confirma actualizar la base de datos de la web?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
                Exit Sub
            End If
        End If
        
        'Inicio conexion a Access
        If Not InicioConexionAccess Then Exit Sub
        
        boCancel = False
        bActualizar.Caption = "&Cancelar"
        Screen.MousePointer = 11
        Ctrl_State True
        act_ProcesoActualizacion
        cAccess.Close
        'cierro la conexión access.
        If chArchivo.Value = 1 Then acc_InvocoArchivo
        
        bActualizar.Caption = "&Inicio"
        Screen.MousePointer = 0
        If bCentinela Then
            Unload Me
            End
            Exit Sub
        End If
        
    ElseIf bActualizar.Caption = "&Cancelar" Then
    
        If MsgBox("¿Confirma cancelar el proceso de actualización?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
            boCancel = True
            bActualizar.Caption = "&Actualizar"
            Ctrl_State False
            Screen.MousePointer = 0
        End If
    Else
        bActualizar.Caption = "&Actualizar"
        Ctrl_State False
    End If
    
End Sub

Private Sub bSalir_Click()
    If bActualizar.Caption = "&Cancelar" Then
        If MsgBox("¿Confirma cancelar la actualización y salir del formulario?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
            boCancel = True
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub bSucesos_Click()
    EjecutarApp sFileErr
End Sub

Private Sub Form_Activate()
    If bload And bCentinela Then
        bload = False
        bActualizar_Click
    Else
        bload = False
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    ObtengoSeteoForm Me, Left, Top
    Me.Height = 4110
    Me.Width = 5820
    bload = True
    Ctrl_State False
    lLogoPaso.Caption = "Definir el tipo de actualización"
    opHacer(1).Value = True
    'x defecto marco todo.
    chEspecie.Value = 1
    chTipo.Value = 1
    chMarca.Value = 1
    chGrupo.Value = 1
    chPlanes.Value = 1
    chGlosario.Value = 1
    chStock.Value = 1
    chPrecio.Value = 1
    chLista.Value = 1
    chAccesorios.Value = 1
    chArchivo.Value = 1
    If Dir(paFileInvoco, vbArchive) = "" Then chArchivo.ForeColor = &HC0&
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    GuardoSeteoForm Me
    CierroConexion
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    End
End Sub

Private Function db_ActualizarUnArticulo() As String
On Error GoTo errAUA

    db_ActualizarUnArticulo = ""
    
    lLogoPaso.Caption = "Lista de Artículos"
    frm_SetDetalleLogo "Copiando artículo a la web"
    
    rdoErrors.Clear
    Cons = "Select * From Articulo " _
                & " Left Outer Join ArticuloFacturacion On ArtID = AFaArticulo " _
                & " Left Outer Join CodigoTexto On Codigo = AFaLista " _
                & " Left Outer Join ArticuloWebPage On ArtID = AWPArticulo " _
            & " Where ArtID = " & Val(tArticulo.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If RsAux("ArtEnUso") Then
            db_ActualizarUnArticulo = acc_InsertoArticulo
        Else
            db_ActualizarUnArticulo = acc_DeleteArticulo(RsAux!ArtID)
        End If
    Else
        db_ActualizarUnArticulo = "Se eliminó el artículo con id " & Val(tArticulo.Tag)
        RsAux.Close
        Exit Function
    End If
    RsAux.Close
    Exit Function
    
errAUA:
    db_ActualizarUnArticulo = f_SetError("Artículo")
End Function

Private Function acc_DeleteArticulo(ByVal lIDArticulo As Long)
On Error GoTo errDA

    'Elimino el artículo.
    rdoErrors.Clear
    Cons = "Select * From ListaArticulos Where LAArtID = " & lIDArticulo
    Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsAcc.EOF Then rsAcc.Delete
    rsAcc.Close
    Exit Function
    
errDA:
    acc_DeleteArticulo = "* " & "Error al eliminar el artículo con id: " & lIDArticulo & vbCrLf & GetRdoError & vbCrLf & Err.Description & vbCrLf
End Function

Private Function acc_InsertoArticulo() As String
'Recibo el resultset con todos los artículos.
Dim sTexto As String
    
    Cons = "Select * From ListaArticulos Where LAArtID = " & RsAux("ArtID")
    Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If rsAcc.EOF Then
        If sArtNew = "" Then sArtNew = RsAux("ArtID") Else sArtNew = sArtNew & "," & RsAux("ArtID")
        rsAcc.AddNew
        rsAcc("LAArtID") = RsAux("ArtID")
    Else
        rsAcc.Edit
    End If
            
    If Trim(sArtMod) <> "" Then sArtMod = sArtMod & ", "
    sArtMod = sArtMod & RsAux("ArtID")
    
    rsAcc("LATipo") = RsAux("ArtTipo")
    rsAcc("LANombre") = RsAux("ArtNombre")
    'NOMBRE WEB
    If Not IsNull(RsAux("AWPNombreArt")) Then
        rsAcc("LANombreWeb") = Trim(RsAux("AWPNombreArt"))
    Else
        rsAcc("LANombreWeb") = Null
    End If
    If Not IsNull(RsAux("AWPAccesorios")) Then
        rsAcc("LAAccesorios") = Trim(RsAux("AWPAccesorios"))
    Else
        rsAcc("LAAccesorios") = Null
    End If
    
    rsAcc("LACodigo") = RsAux("ArtCodigo")
    rsAcc("LAEnWeb") = RsAux("ArtEnWeb")
    If Not IsNull(RsAux("Clase")) Then
        rsAcc("LALista") = RsAux("Clase")
    Else
        rsAcc("LALista") = Null
    End If
    If Not IsNull(RsAux("AWPSinPrecio")) Then
        rsAcc("LASinPrecio") = RsAux("AWPSinPrecio")
    Else
        rsAcc("LASinPrecio") = False
    End If
    
    If Not IsNull(RsAux!AWPFoto) Then
        If Trim(RsAux!AWPFoto) = "" Then
            rsAcc("LAFoto") = Null
        Else
            rsAcc("LAFoto") = Trim(RsAux!AWPFoto)
        End If
    Else
        rsAcc("LAFoto") = Null
    End If
    
    sTexto = RetornoLongText("AWPTexto")
    If sTexto <> "" Then
        rsAcc("LATextoHTML") = sTexto
    Else
        rsAcc("LATextoHTML") = Null
    End If
    sTexto = RetornoLongText("AWPTexto2")
    If sTexto <> "" Then
        rsAcc("LATextoInterno") = sTexto
    Else
        rsAcc("LATextoInterno") = Null
    End If
    rsAcc("LAEsCombo") = RsAux("ArtEsCombo")
    
    'Lo pongo en False y si no es así lo cambio.
    rsAcc("LAHabilitado") = False
    If Not IsNull(RsAux("ArtHabilitado")) Then
        If UCase(RsAux("ArtHabilitado")) = "S" Then
            rsAcc("LAHabilitado") = True
        End If
    End If
    rsAcc.Update
    rsAcc.Close
    Exit Function
errIA:
    acc_InsertoArticulo = "* " & "Error al agregar o modificar el artículo con id: " & RsAux!ArtID & vbCrLf & GetRdoError & vbCrLf & Err.Description & vbCrLf
End Function

Private Function db_ArticulosModificados() As Integer
On Error GoTo errAM
    
    db_ArticulosModificados = 0
    rdoErrors.Clear
    Cons = "Select Count(*) From Articulo " _
        & " Where ArtModificado >= '" & Format(paFUltActualizacion, "mm/dd/yyyy hh:nn:ss") & "'" _
        & " And (ArtID IN " _
            & "(Select Distinct(PViArticulo) from PrecioVigente Where PViMoneda = 1 And PViHabilitado = 1) " _
            & " Or ArtEsCombo = 1)"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    rdoErrors.Clear
    If RsAux(0) > 0 Then db_ArticulosModificados = RsAux(0)
    RsAux.Close
    Exit Function
    
errAM:
    MsgBox "Error"
End Function

Private Sub act_ProcesoActualizacion()
Dim sError As String
Dim sUltEjec As String

    Set colError = New Collection
    tSuceso.Text = ""
    
    lCantArtMod = 0
    If Not IsDate(paFUltActualizacion) Then paFUltActualizacion = DateAdd("d", -15, Now)
        
    sArtMod = ""
    If Val(tArticulo.Tag) = 0 Then
        If opHacer(1).Value Then
            'Busco los artículos que fueron modificados.
            lCantArtMod = db_ArticulosModificados
        End If
    Else
        lCantArtMod = 1
    End If
    'Paso las especies.
    
    sError = acc_ActualizarEspecie
    If sError <> "" Then
        colError.Add sError
        tSuceso.Text = tSuceso.Text & sError
    End If
    DoEvents
    If boCancel Then
        pbParcial.Value = 0
        Exit Sub
    End If
    
    sError = acc_ActualizarTipoArticulo
    If sError <> "" Then colError.Add sError: tSuceso.Text = tSuceso.Text & sError
    DoEvents
    If boCancel Then
        pbParcial.Value = 0
        Exit Sub
    End If
    
    sError = acc_ActualizarMarca
    If sError <> "" Then colError.Add sError: tSuceso.Text = tSuceso.Text & sError
    DoEvents
    If boCancel Then
        pbParcial.Value = 0
        Exit Sub
    End If
    
    sError = acc_ActualizarGrupoyArticuloGrupo
    If sError <> "" Then colError.Add sError: tSuceso.Text = tSuceso.Text & sError
    DoEvents
    If boCancel Then
        pbParcial.Value = 0
        Exit Sub
    End If
        
    If Val(tArticulo.Tag) > 0 Then
        sError = db_ActualizarUnArticulo
    Else
        sError = acc_ActualizarArticulo
    End If
    If Trim(sArtMod) = "" Then sArtMod = "0"
    If sError <> "" Then colError.Add sError: tSuceso.Text = tSuceso.Text & sError
    DoEvents
    If boCancel Then
        pbParcial.Value = 0
        Exit Sub
    End If
    
    sError = acc_ActualizarTipoCuota
    If sError <> "" Then colError.Add sError: tSuceso.Text = tSuceso.Text & sError
    DoEvents
    If boCancel Then
        pbParcial.Value = 0
        Exit Sub
    End If
    
    sError = acc_ActualizarTipoPlan
    If sError <> "" Then colError.Add sError: tSuceso.Text = tSuceso.Text & sError
    DoEvents
    If boCancel Then
        pbParcial.Value = 0
        Exit Sub
    End If

    sError = acc_PasoCoeficiente
    If sError <> "" Then colError.Add sError: tSuceso.Text = tSuceso.Text & sError
    DoEvents
    If boCancel Then
        pbParcial.Value = 0
        Exit Sub
    End If
    
    
    sError = acc_OcultoArticulo
    If sError <> "" Then colError.Add sError: tSuceso.Text = tSuceso.Text & sError
    DoEvents
    If boCancel Then
        pbParcial.Value = 0
        Exit Sub
    End If
    
    sError = PasoPrecioVigente
    If sError <> "" Then colError.Add sError: tSuceso.Text = tSuceso.Text & sError
    sError = acc_PasoPreciosACombo
    If sError <> "" Then colError.Add sError: tSuceso.Text = tSuceso.Text & sError
    DoEvents
    If boCancel Then
        pbParcial.Value = 0
        Exit Sub
    End If
    
    sError = acc_OcultoTipoArticulo
    If sError <> "" Then colError.Add sError: tSuceso.Text = tSuceso.Text & sError
    DoEvents
    If boCancel Then
        pbParcial.Value = 0
        Exit Sub
    End If

    sError = acc_OcultoEspecie
    If sError <> "" Then colError.Add sError: tSuceso.Text = tSuceso.Text & sError
    DoEvents
    If boCancel Then
        pbParcial.Value = 0
        Exit Sub
    End If
        
    sError = acc_ActualizarGlosario
    If sError <> "" Then colError.Add sError: tSuceso.Text = tSuceso.Text & sError
    DoEvents
    If boCancel Then
        pbParcial.Value = 0
        Exit Sub
    End If
    
    sError = acc_ActualizoStock
    If sError <> "" Then colError.Add sError: tSuceso.Text = tSuceso.Text & sError
    DoEvents
    If boCancel Then
        pbParcial.Value = 0
        Exit Sub
    End If
    
    'Me guardo la ejecución anterior antes de actualizar
    sUltEjec = paFUltActualizacion
    If Val(tArticulo.Tag) = 0 Then
        sError = db_SetUltimaEjecucion
        If sError <> "" Then colError.Add sError: tSuceso.Text = tSuceso.Text & sError
        DoEvents
        If boCancel Then
            pbParcial.Value = 0
            Exit Sub
        End If
    End If
        
    sError = acc_ActualizarArticuloAccesorios
    If sError <> "" Then colError.Add sError: tSuceso.Text = tSuceso.Text & sError
    DoEvents
    If boCancel Then
        pbParcial.Value = 0
        Exit Sub
    End If
    
    'Esta va si o si ya que no se
    sError = acc_ActualizarListasDePrecios
    If sError <> "" Then colError.Add sError: tSuceso.Text = tSuceso.Text & sError
    DoEvents
    If boCancel Then
        pbParcial.Value = 0
        Exit Sub
    End If
        
    sError = GeneroPaginasDeArticulos(sUltEjec)
    If sError <> "" Then colError.Add sError: tSuceso.Text = tSuceso.Text & sError
    DoEvents
    If boCancel Then
        pbParcial.Value = 0
        Exit Sub
    End If
    pbParcial.Value = 0
    
    If colError.Count > 0 And bCentinela Then GraboErrores
    
    lLogoPaso.Caption = "Web Actualizada"
    frm_SetDetalleLogo ""
    
End Sub

Private Sub Ctrl_State(ByVal bRun As Boolean)
    
    frmOpcion.Visible = Not bRun
            
    tSuceso.Visible = bRun
    pbParcial.Visible = bRun
    If bRun Then tSuceso.Move 120, 840
    
End Sub

Private Function acc_ActualizarEspecie() As String
On Error GoTo errPE
Dim lCant As Long, lSum As Long
Dim sEspecieEnCGSA As String

    acc_ActualizarEspecie = ""
    If Not (chEspecie.Value = 1 Or opHacer(0).Value) Then Exit Function
    
    lLogoPaso.Caption = "Copiando especies."
    lLogoDet.Caption = "Validando ..."
    
    sEspecieEnCGSA = ""
    lCant = 0
    
    'Saco la cantidad de especies que hay en la base de datos.
    Cons = "Select Count(*), Sum(Len(RTrim(EspNombre))) From Especie"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux(0) > 0 Then lCant = RsAux(0): lSum = RsAux(1)
    RsAux.Close
    
    If lCant > 0 Then
        'Access si coinciden ambas me voy.
        Set RsAux = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        'Si no me cambio la cantidad ni la suma de los strings.
        If RsAux(0) = lCant And lSum = RsAux(1) Then
            lCant = 0
        End If
        RsAux.Close
    Else
        frm_SetDetalleLogo "Eliminando ..."
        Cons = "Delete * From Especie"
        cAccess.Execute (Cons)
    End If
    
    If lCant > 0 Then
        frm_SetDetalleLogo "Pasando datos ..."
        
        pbParcial.Max = lCant + 1       'Le sumo 1 porque borro abajo.
        Cons = "Select * From Especie"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            
            'Veo si la especie existe, sino la inserto.
            Cons = "Select * From Especie Where EspCodigo = " & RsAux("EspCodigo")
            Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If rsAcc.EOF Then
                rsAcc.AddNew
                rsAcc("EspMostrar") = False
            Else
                rsAcc.Edit
            End If
            rsAcc("EspCodigo") = RsAux("EspCodigo")
            rsAcc("EspNombre") = RsAux("EspNombre")
            rsAcc.Update
            rsAcc.Close
            
            If sEspecieEnCGSA <> "" Then sEspecieEnCGSA = sEspecieEnCGSA & ", "
            sEspecieEnCGSA = sEspecieEnCGSA & RsAux("EspCodigo")
            
            RsAux.MoveNext
            AumentoProgressParcial
            
        Loop
        RsAux.Close
        
        frm_SetDetalleLogo "Eliminando ..."
        Cons = "Delete * From Especie Where EspCodigo Not In (" & sEspecieEnCGSA & ")"
        cAccess.Execute (Cons)
        
        AumentoProgressParcial
    End If
    
    pbParcial.Value = 0
    frm_SetDetalleLogo ""
    Exit Function
    
errPE:
    acc_ActualizarEspecie = f_SetError("Especies")
    pbParcial.Value = 0
End Function

Private Function acc_ActualizarTipoArticulo() As String
On Error GoTo errPTA
Dim lCant As Long, lSum As Long
    
    acc_ActualizarTipoArticulo = ""
    If Not (chTipo.Value = 1 Or opHacer(0).Value) Then Exit Function
    
    lLogoPaso.Caption = "Tipos de Artículos"
    frm_SetDetalleLogo "Validando Tipos de Artículos ..."
    
    sTipoEnCGSA = ""
    
    Cons = "Select Count(*), Sum(Len(rTrim(TipNombre))) From Tipo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux(0) > 0 Then lCant = RsAux(0): lSum = RsAux(1)
    RsAux.Close
    
    If lCant > 0 Then
        'Comparo con la de access.
        Set RsAux = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux(0) = lCant And lSum = RsAux(1) Then lCant = 0
        RsAux.Close
    Else
        frm_SetDetalleLogo "Eliminando ..."
        Cons = "Delete * From Tipo"
        cAccess.Execute (Cons)
    End If
    
    If lCant > 0 Then
        
        frm_SetDetalleLogo "Pasando datos ..."
        
        pbParcial.Max = lCant + 1       'Le sumo 1 porque borro abajo.
        
        Cons = "Select * From Tipo"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
                    
            'Veo si la especie existe, sino la inserto.
            Cons = "Select * From Tipo Where TipCodigo = " & RsAux("TipCodigo")
            Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If rsAcc.EOF Then
                rsAcc.AddNew
                rsAcc("TipCodigo") = RsAux("TipCodigo")
                rsAcc("TipMostrar") = False
            Else
                rsAcc.Edit
            End If
            rsAcc("TipNombre") = RsAux("TipNombre")
            rsAcc("TipEspecie") = RsAux("TipEspecie")
            If Not IsNull(RsAux("TipAbreviacion")) Then
                rsAcc("TipAbreviacion") = RsAux("TipAbreviacion")
            Else
                rsAcc("TipAbreviacion") = Null
            End If
            If Not IsNull(RsAux("TipLocalRep")) Then
                rsAcc("TipLocalRep") = RsAux("TipLocalRep")
            Else
                rsAcc("TipLocalRep") = Null
            End If
            If Not IsNull(RsAux("TipBusqWeb")) Then
                rsAcc("TipBusqWeb") = RsAux("TipBusqWeb")
            Else
                rsAcc("TipBusqWeb") = Null
            End If
            If Not IsNull(RsAux("TipArrayCaract")) Then
                rsAcc("TipArrayCaract") = RsAux("TipArrayCaract")
            Else
                rsAcc("TipArrayCaract") = Null
            End If
            rsAcc.Update
            rsAcc.Close
            
            If sTipoEnCGSA <> "" Then sTipoEnCGSA = sTipoEnCGSA & ", "
            sTipoEnCGSA = sTipoEnCGSA & RsAux("TipCodigo")
            
            RsAux.MoveNext
            AumentoProgressParcial
        Loop
        RsAux.Close
        
        frm_SetDetalleLogo "Eliminando ..."
        'Elimino los tipos que fueron eliminados en CGSA.
        Cons = "Delete * From Tipo Where TipCodigo Not In ( " & sTipoEnCGSA & ")"
        cAccess.Execute (Cons)
        AumentoProgressParcial
        
    End If
    
    frm_SetDetalleLogo ""
    pbParcial.Value = 0
    Exit Function
    
errPTA:
    acc_ActualizarTipoArticulo = f_SetError("Tipos de Artículos")
    pbParcial.Value = 0
End Function

Private Function InsertoArticuloPorPrecio(ByVal idArt As Long) As Boolean
Dim bInsertar As Boolean
Dim sTexto As String
Dim rsArAcc As rdoResultset
    
    InsertoArticuloPorPrecio = False
    Cons = "Select * From Articulo " _
            & " Left Outer Join ArticuloFacturacion On ArtID = AFaArticulo " _
                & " Left Outer Join CodigoTexto On Codigo = AFaLista " _
            & " Left Outer Join ArticuloWebPage On ArtID = AWPArticulo " _
        & " Where ArtID = " & idArt
        
    Set rsAr = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If rsAr.EOF Then
        rsAr.Close
        Exit Function
        
    Else
        
        If rsAr("ArtEnUso") Then
            Cons = "Select * From ListaArticulos Where LAArtID = " & rsAr("ArtID")
            Set rsArAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
            If rsArAcc.EOF Then
                
                rsArAcc.AddNew
                rsArAcc("LAArtID") = rsAr("ArtID")
                
                If Trim(sArtMod) <> "" Then sArtMod = sArtMod & ", "
                sArtMod = sArtMod & rsAr("ArtID")
                
                rsArAcc("LATipo") = rsAr("ArtTipo")
                rsArAcc("LANombre") = rsAr("ArtNombre")
                rsArAcc("LACodigo") = rsAr("ArtCodigo")
                rsArAcc("LAEnWeb") = rsAr("ArtEnWeb")
                
                If Not IsNull(rsAr!AWPFoto) Then
                    If Trim(rsAr!AWPFoto) = "" Then
                        rsArAcc("LAFoto") = Null
                    Else
                        rsArAcc("LAFoto") = rsAr!AWPFoto
                    End If
                Else
                    rsArAcc("LAFoto") = Null
                End If
                
                If Not IsNull(rsAr("Clase")) Then
                    rsArAcc("LALista") = rsAr("Clase")
                Else
                    rsArAcc("LALista") = Null
                End If
                If Not IsNull(rsAr("AWPSinPrecio")) Then
                    rsArAcc("LASinPrecio") = rsAr("AWPSinPrecio")
                Else
                    rsArAcc("LASinPrecio") = False
                End If
                sTexto = RetornoLongText2("AWPTexto")
                If sTexto <> "" Then
                    rsArAcc("LATextoHTML") = sTexto
                Else
                    rsArAcc("LATextoHTML") = Null
                End If
                sTexto = RetornoLongText2("AWPTexto2")
                If sTexto <> "" Then
                    rsArAcc("LATextoInterno") = sTexto
                Else
                    rsArAcc("LATextoInterno") = Null
                End If
                rsArAcc("LAEsCombo") = rsAr("ArtEsCombo")
                
                'Lo pongo en False y si no es así lo cambio.
                rsArAcc("LAHabilitado") = False
                
                If Not IsNull(rsAr("ArtHabilitado")) Then
                    If UCase(rsAr("ArtHabilitado")) = "S" Then
                        rsArAcc("LAHabilitado") = True
                    End If
                End If
                rsArAcc.Update
            End If
            rsArAcc.Close
        End If
        rsAr.Close
    End If
    
End Function

Private Function acc_ActualizarArticulo() As String
On Error GoTo errPA
Dim bInsertar As Boolean
Dim lCant As Long
Dim sTexto As String

'Busco los artículos que estén en uso y tengan
'fecha de modificación mayor o = a la última actualización.

    acc_ActualizarArticulo = ""
    lLogoPaso.Caption = "Artículos"
    sArtMod = " "
    lCant = lCantArtMod
    
    If opHacer(0).Value = True Then
        
        frm_SetDetalleLogo "Pasando todos los artículos."
        
        Cons = "Select Count(*) From Articulo " _
            & " Left Outer Join ArticuloFacturacion On ArtID = AFaArticulo " _
                & " Left Outer Join CodigoTexto On Codigo = AFaLista " _
            & " Left Outer Join ArticuloWebPage On ArtID = AWPArticulo " _
            & " Where (ArtID IN " _
                & "(Select Distinct(PViArticulo) from PrecioVigente Where PViMoneda = 1 And PViHabilitado = 1) Or ArtEsCombo = 1)"
        
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        lCant = RsAux(0)
        RsAux.Close
        
        Cons = "Select * From Articulo " _
            & " Left Outer Join ArticuloFacturacion On ArtID = AFaArticulo " _
                & " Left Outer Join CodigoTexto On Codigo = AFaLista " _
                & " Left Outer Join ArticuloWebPage On ArtID = AWPArticulo " _
            & " Where (ArtID IN " _
                & "(Select Distinct(PViArticulo) from PrecioVigente Where PViMoneda = 1 And PViHabilitado = 1) Or ArtEsCombo = 1)"
        
    Else
    
        If lCantArtMod = 0 Then
            
            GoTo evSalir
            
        Else
            
            frm_SetDetalleLogo "Pasando artículos modificados."
            
            Cons = "Select * From Articulo " _
                & " Left Outer Join ArticuloFacturacion On ArtID = AFaArticulo " _
                    & " Left Outer Join CodigoTexto On Codigo = AFaLista " _
                & " Left Outer Join ArticuloWebPage On ArtID = AWPArticulo " _
                & " Where ArtModificado >= '" & Format(paFUltActualizacion, "mm/dd/yyyy hh:nn:ss") & "'" _
                & " And (ArtID IN " _
                    & "(Select Distinct(PViArticulo) from PrecioVigente Where PViMoneda = 1 And PViHabilitado = 1) Or ArtEsCombo = 1)"
                
        End If
        
    End If
        
    pbParcial.Max = lCant
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
    Do While Not RsAux.EOF
        If RsAux("ArtEnUso") Then
            acc_ActualizarArticulo = acc_ActualizarArticulo & acc_InsertoArticulo
        Else
            acc_ActualizarArticulo = acc_ActualizarArticulo & acc_DeleteArticulo(RsAux!ArtID)
        End If
        RsAux.MoveNext
        AumentoProgressParcial
    Loop
    RsAux.Close
    
evSalir:
    pbParcial.Value = 0
    Exit Function
    
errPA:
    acc_ActualizarArticulo = f_SetError("Articulos")
    pbParcial.Value = 0
End Function
Private Function RetornoLongText2(ByVal sCampo As String) As String
On Error GoTo errRLT
    RetornoLongText2 = ""
    RetornoLongText2 = rsAr(sCampo)
errRLT:
End Function

Private Function RetornoLongText(ByVal sCampo As String) As String
On Error GoTo errRLT
    RetornoLongText = ""
    RetornoLongText = RsAux(sCampo)
errRLT:
End Function

Private Function acc_OcultoTipoArticulo() As String
On Error GoTo errET

    acc_OcultoTipoArticulo = ""
    
    lLogoPaso.Caption = "Ocultando Tipos"
    frm_SetDetalleLogo "Validando ..."
    
    Cons = "Select Count(*) From Tipo Where TipCodigo Not In (Select Distinct(LATipo) From ListaArticulos)"
    Set RsAux = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux(0) Then
        RsAux.Close
        Exit Function
    Else
        pbParcial.Max = RsAux(0)
        RsAux.Close
    End If
    
    Cons = "Select * From Tipo Where TipCodigo Not In (Select Distinct(LATipo) From ListaArticulos)"
    Set RsAux = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        RsAux.Edit
        RsAux("TipMostrar") = False
        RsAux.Update
        RsAux.MoveNext
        AumentoProgressParcial
    Loop
    RsAux.Close
    frm_SetDetalleLogo ""
    pbParcial.Value = 0
    Exit Function
    
errET:
    acc_OcultoTipoArticulo = f_SetError("Oculto Tipos")
    pbParcial.Value = 0
End Function

Private Function acc_OcultoEspecie() As String
On Error GoTo errEE

    acc_OcultoEspecie = ""
    
    lLogoPaso.Caption = "Ocultando Especies"
    frm_SetDetalleLogo "Validando ..."
    
    'Marco para que no muestre las especies que no tienen artículos.
    Cons = "Select * From Especie Where EspCodigo Not In (Select Distinct(TipEspecie) From Tipo) And EspMostrar <> 0" _
        & " Union All " _
        & " Select * From Especie Where EspCodigo In (Select Distinct(TipEspecie) From Tipo Where TipMostrar = 0) And EspMostrar <> 0"
    
    Set RsAux = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        RsAux.Edit
        RsAux("EspMostrar") = False
        RsAux.Update
        RsAux.MoveNext
    Loop
    RsAux.Close
    frm_SetDetalleLogo ""
    Exit Function
    
errEE:
    acc_OcultoEspecie = f_SetError("Oculto Especies")
    pbParcial.Value = 0
End Function

Private Function acc_OcultoArticulo() As String
On Error GoTo errOA
Dim lCant As Long
Dim sArtEnAcc As String

    acc_OcultoArticulo = ""
    lLogoPaso.Caption = "Ocultando artículos"
    frm_SetDetalleLogo "Validando ..."
    
    pbParcial.Value = 0
    pbParcial.Max = 8
    pbParcial.Value = 1
    '---------------------------------------------------------------------------
    
    'Veo los que están en access pero no es CGSA.
    sArtEnAcc = ""
    rdoErrors.Clear
    
    Cons = "Select * From Articulo Where ArtEnUso = 1"
    Set rsAcc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    rdoErrors.Clear
    AumentoProgressParcial
    
    Do While Not rsAcc.EOF
        If sArtEnAcc <> "" Then sArtEnAcc = sArtEnAcc & ", "
        sArtEnAcc = sArtEnAcc & rsAcc("ArtID")
        rsAcc.MoveNext
    Loop
    rsAcc.Close
    
    AumentoProgressParcial
    
    rdoErrors.Clear
    
    'Ahora consulto los que esten en access y no esten en CGSA.
    Cons = "Delete * From ListaArticulos Where LAArtID Not IN (" & sArtEnAcc & ")"
    cAccess.Execute (Cons)
    
    
    pbParcial.Value = pbParcial.Value + 2
    '---------------------------------------------------------------------------
    frm_SetDetalleLogo "Validando precios..."
    
    rdoErrors.Clear
    'Elimino Los precios de los artículos que se eliminaron.
    Cons = "Delete * From PrecioVigente Where PViArticulo Not In " _
        & " (Select Distinct(LAArtID) From ListaArticulos)"
    cAccess.Execute (Cons)
    
    rdoErrors.Clear
    
    pbParcial.Value = pbParcial.Value + 2
    pbParcial.Value = 0
    Exit Function
    
errOA:
    acc_OcultoArticulo = "(OcultoArticulo) " & Trim(Err.Description)
    pbParcial.Value = 0
End Function

Private Function acc_PasoCoeficiente() As String
On Error GoTo errPC
Dim lSum As Long
Dim cSum As Currency

    acc_PasoCoeficiente = ""
    
    lLogoPaso.Caption = "Coeficientes"
    frm_SetDetalleLogo "Validando ..."
    
    rdoErrors.Clear
    Cons = "Select Count(*), Sum(CoePlan) + Sum(CoeCoeficiente) From Coeficiente"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    rdoErrors.Clear
    If RsAux(0) = 0 Then
        RsAux.Close
        rdoErrors.Clear
        Cons = "Delete * From Coeficiente"
        cAccess.Execute (Cons)
        Exit Function
    Else
        pbParcial.Max = RsAux(0) + 2
        cSum = RsAux(1)
        RsAux.Close
    End If
    
    rdoErrors.Clear
    Set RsAux = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    rdoErrors.Clear
    If RsAux(0) = pbParcial.Max - 2 And cSum = RsAux(1) Then
        RsAux.Close
        pbParcial.Value = 0
        Exit Function
    Else
        RsAux.Close
    End If
    
    
    pbParcial.Value = 1
    '1ero Elimino todos los que tengo en Access.
    frm_SetDetalleLogo "Pasando ..."
    rdoErrors.Clear
    Cons = "Delete * From Coeficiente"
    cAccess.Execute (Cons)
    pbParcial.Value = 2
    
    rdoErrors.Clear
    Cons = "Select * From Coeficiente"
    Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    rdoErrors.Clear
    
    Cons = "Select * From Coeficiente"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    rdoErrors.Clear
    
    Do While Not RsAux.EOF
        
        rsAcc.AddNew
        rsAcc("CoePlan") = RsAux("CoePlan")
        rsAcc("CoeTipoCuota") = RsAux("CoeTipoCuota")
        rsAcc("CoeMoneda") = RsAux("CoeMoneda")
        rsAcc("CoeCoeficiente") = RsAux("CoeCoeficiente")
        rsAcc.Update
        
        RsAux.MoveNext
        AumentoProgressParcial
    Loop
    RsAux.Close
    pbParcial.Value = 0
    Exit Function
    
errPC:
    acc_PasoCoeficiente = f_SetError("Coeficientes")
    pbParcial.Value = 0
End Function

Private Function acc_ActualizarTipoPlan()
On Error GoTo errPTP
Dim lCant As Long, cSum As Currency

    acc_ActualizarTipoPlan = ""
    If Not (chPlanes.Value = 1 Or opHacer(0).Value) Then Exit Function
    
    lLogoPaso.Caption = "Tipos de Planes"
    frm_SetDetalleLogo "Validando ..."
    
    Cons = "Select Count(*), Sum(len(rtrim(PlaNombre))) + Sum(len(rtrim(PlaDetalle))) + Sum(PlaCoeficiente) + Sum(PlaCodigo) From TipoPlan"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux(0) = 0 Then
        RsAux.Close
        'Elimino las del access.
        frm_SetDetalleLogo "Eliminando ..."
        Cons = "Delete * From TipoPlan"
        cAccess.Execute (Cons)
        Exit Function
    Else
        lCant = RsAux(0)
        cSum = RsAux(1)
        pbParcial.Max = RsAux(0) + 2
        RsAux.Close
    End If
    
    'Válido contra la tabla de acces.
    Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If rsAcc(0) = lCant And cSum = rsAcc(1) Then lCant = 0
    rsAcc.Close
    
    If lCant > 0 Then
        pbParcial.Value = 1
        
        frm_SetDetalleLogo "Pasando datos ..."
        
        Cons = "Delete * From TipoPlan"
        cAccess.Execute (Cons)
        
        pbParcial.Value = 2
        
        Cons = "Select * From TipoPlan"
        Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
        Cons = "Select * From TipoPlan"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            
            rsAcc.AddNew
            rsAcc("PlaCodigo") = RsAux("PlaCodigo")
            rsAcc("PlaNombre") = RsAux("PlaNombre")
            rsAcc("PlaCoeficiente") = RsAux("PlaCoeficiente")
            If Not IsNull(RsAux("PlaDetalle")) Then
                rsAcc("PlaDetalle") = RsAux("PlaDetalle")
            Else
                rsAcc("PlaDetalle") = Null
            End If
            If Not IsNull(RsAux("PlaCreado")) Then
                rsAcc("PlaCreado") = RsAux("PlaCreado")
            Else
                rsAcc("PlaCreado") = Null
            End If
            rsAcc.Update
            
            RsAux.MoveNext
            AumentoProgressParcial
        Loop
        RsAux.Close
        
        rsAcc.Close
    End If
    pbParcial.Value = 0
    Exit Function
    
errPTP:
    acc_ActualizarTipoPlan = f_SetError("Tipos de Planes")
    pbParcial.Value = 0
    
End Function

Private Function acc_ActualizarTipoCuota() As String
On Error GoTo errPTC
Dim lCant As Long, lSum As Long
Dim sTipoCuotaEnCGSA As String

    acc_ActualizarTipoCuota = ""
    If Not (chPlanes.Value = 1 Or opHacer(0).Value) Then Exit Function
    
    lLogoPaso.Caption = "Tipos de Cuotas"
    frm_SetDetalleLogo "Validando Tipo de Cuotas ..."
    
    'Sumo el nombre + la abreviación.
    Cons = "Select Count(*), Sum(Len(rTrim(TcuNombre))+Len(rTrim(TcuAbreviacion))) From TipoCuota"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux(0) > 0 Then lCant = RsAux(0): lSum = RsAux(1)
    RsAux.Close
    
    If lCant > 0 Then
        Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If rsAcc(0) = lCant And lSum = rsAcc(1) Then lCant = 0
        rsAcc.Close
    Else
        frm_SetDetalleLogo " Eliminando ..."
        Cons = "Delete * From TipoCuota"
        cAccess.Execute (Cons)
    End If
    
    If lCant > 0 Then
        
        frm_SetDetalleLogo "Pasando datos ..."
        
        pbParcial.Max = lCant + 4       'Le sumo porque borro abajo.
        
        lCant = 0
        Cons = "Select * From TipoCuota"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF

            Cons = "Select * From TipoCuota Where TCuCodigo = " & RsAux("TCuCodigo")
            Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If rsAcc.EOF Then
                rsAcc.AddNew
            Else
                rsAcc.Edit
            End If
            rsAcc("TCuCodigo") = RsAux("TCuCodigo")
            rsAcc("TCuNombre") = RsAux("TCuNombre")
            rsAcc("TCuAbreviacion") = RsAux("TCuAbreviacion")
            rsAcc("TCuCantidad") = RsAux("TCuCantidad")
            If Not IsNull(RsAux("TCuVencimientoE")) Then
                rsAcc("TCuVencimientoE") = RsAux("TCuVencimientoE")
            Else
                rsAcc("TCuVencimientoE") = Null
            End If
            rsAcc("TCuVencimientoC") = RsAux("TCuVencimientoC")
            If Not IsNull(RsAux("TCuDificultad")) Then
                rsAcc("TCuDificultad") = RsAux("TCuDificultad")
            Else
                rsAcc("TCuDificultad") = Null
            End If
            If Not IsNull(RsAux("TCuDistancia")) Then
                rsAcc("TCuDistancia") = RsAux("TCuDistancia")
            Else
                rsAcc("TCuDistancia") = Null
            End If
            If Not IsNull(RsAux("TCuDeshabilitado")) Then
                rsAcc("TCuDeshabilitado") = RsAux("TCuDeshabilitado")
            Else
                rsAcc("TCuDeshabilitado") = Null
            End If
            rsAcc("TCuEspecial") = RsAux("TCuEspecial")
            If Not IsNull(RsAux("TCuOrden")) Then
                rsAcc("TCuOrden") = RsAux("TCuOrden")
            Else
                rsAcc("TCuOrden") = Null
            End If
            rsAcc.Update
            rsAcc.Close
            
            If sTipoCuotaEnCGSA <> "" Then sTipoCuotaEnCGSA = sTipoCuotaEnCGSA & ", "
            sTipoCuotaEnCGSA = sTipoCuotaEnCGSA & RsAux("TCuCodigo")
            
            RsAux.MoveNext
            AumentoProgressParcial
        Loop
        RsAux.Close
        
        frm_SetDetalleLogo "Eliminando ..."
        Cons = "Delete * From TipoCuota Where TCuCodigo Not In (" & sTipoCuotaEnCGSA & ")"
        cAccess.Execute (Cons)
        pbParcial.Value = pbParcial.Value + 2
       
        'Elmino los Precios que tengan un tipo de cuota que no exista más.
        Cons = "Delete * From PrecioVigente Where PViTipoCuota Not In (" & sTipoCuotaEnCGSA & ")"
        cAccess.Execute (Cons)
    End If
    
    frm_SetDetalleLogo ""
    pbParcial.Value = 0
    Exit Function
    
errPTC:
    acc_ActualizarTipoCuota = f_SetError("Tipos de cuotas")
    pbParcial.Value = 0
End Function

Private Function acc_ActualizarGrupoyArticuloGrupo() As String
On Error GoTo errPM
Dim lCant As Long, lSum As Long
Dim sDel As String
    
    acc_ActualizarGrupoyArticuloGrupo = ""
    If Not (chGrupo.Value = 1 Or opHacer(0).Value) Then Exit Function
    
    lLogoPaso.Caption = "Grupos"
    frm_SetDetalleLogo "Validando Grupos ..."
    sDel = ""
    
    lCant = 0
    'Saco la cantidad de especies que hay en la base de datos.
    Cons = "Select Count(*), Sum(Len(rTrim(GruNombre))) From Grupo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux(0) > 0 Then lCant = RsAux(0): lSum = RsAux(1)
    RsAux.Close
    
    Set RsAux = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux(0) = lCant And RsAux(1) = lSum Then
        lCant = 0
    End If
    RsAux.Close
    
    If lCant > 0 Then
        frm_SetDetalleLogo "Pasando datos ..."
        pbParcial.Max = lCant + 1       'Le sumo 1 porque borro abajo.
        
        Cons = "Select * From Grupo"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            
            'Veo si la especie existe, sino la inserto.
            Cons = "Select * From Grupo Where GruCodigo = " & RsAux("GruCodigo")
            Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If rsAcc.EOF Then
                rsAcc.AddNew
                rsAcc("GruCodigo") = RsAux("GruCodigo")
            Else
                rsAcc.Edit
            End If
            rsAcc("GruNombre") = RsAux("GruNombre")
            rsAcc.Update
            rsAcc.Close
            
            If sDel <> "" Then sDel = sDel & ", "
            sDel = sDel & RsAux("GruCodigo")
            
            RsAux.MoveNext
            AumentoProgressParcial
        Loop
        RsAux.Close
        
        frm_SetDetalleLogo "Eliminando ..."
        Cons = "Delete * From Grupo Where GruCodigo Not In (" & sDel & ")"
        cAccess.Execute (Cons)
        AumentoProgressParcial
        
    End If
    
    pbParcial.Value = 0
    
    lLogoPaso.Caption = "Grupos de Artículos"
    frm_SetDetalleLogo "Validando ArticuloGrupo ..."
        
    lCant = 0
    'Saco la cantidad de ArtGrupo que hay en la base de datos.
    Cons = "Select Count(*), Sum(AGrGrupo) + Sum(AGrArticulo) From ArticuloGrupo Where AGrArticulo In (Select ArtID From Articulo Where ArtEnUso = 1)"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux(0) > 0 Then lCant = RsAux(0): lSum = RsAux(1)
    RsAux.Close
    
    If lCant = 0 Then
        frm_SetDetalleLogo "Eliminando ..."
        Cons = "Delete * From ArticuloGrupo"
        cAccess.Execute (Cons)
        Exit Function
    End If
    
    Cons = "Select Count(*), Sum(AGrGrupo) + Sum(AGrArticulo) From ArticuloGrupo"
    Set RsAux = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux(0) = lCant And RsAux(1) = lSum Then lCant = 0
    RsAux.Close
    
    If lCant > 0 Then
        frm_SetDetalleLogo "Copiando ArticuloGrupo ..."
        pbParcial.Max = lCant + 1       'Le sumo 1 porque borro abajo.
        
        Cons = "Delete * From ArticuloGrupo"
        cAccess.Execute (Cons)
        
        pbParcial.Value = 1
        
        Cons = "Select * From ArticuloGrupo"
        Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        Cons = "Select * From ArticuloGrupo Where AGrArticulo In (Select ArtID From Articulo Where ArtEnUso = 1)"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            
            'Veo si la especie existe, sino la inserto.
            rsAcc.AddNew
            rsAcc("AGrArticulo") = RsAux("AGrArticulo")
            rsAcc("AGrGrupo") = RsAux("AGrGrupo")
            rsAcc.Update
            
            RsAux.MoveNext
            AumentoProgressParcial
        Loop
        RsAux.Close
        rsAcc.Close
    End If
    
    frm_SetDetalleLogo ""
    pbParcial.Value = 0
    Exit Function
errPM:
    acc_ActualizarGrupoyArticuloGrupo = f_SetError("Grupos de Artículos")
    pbParcial.Value = 0
End Function
Private Function acc_ActualizarMarca() As String
On Error GoTo errPM
Dim lCant As Long, lSum As Long
Dim sMarca As String
    
    acc_ActualizarMarca = ""
    If Not (chMarca.Value = 1 Or opHacer(0).Value) Then Exit Function
    
    lLogoPaso.Caption = "Pasando Marcas de artículos"
    frm_SetDetalleLogo "Validando Marcas ..."
    sMarca = ""
    
    lCant = 0
    'Saco la cantidad de especies que hay en la base de datos.
    Cons = "Select Count(*), Sum(Len(rTrim(MarNombre))) From Marca"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux(0) > 0 Then lCant = RsAux(0): lSum = RsAux(1)
    RsAux.Close
    If lCant > 0 Then
        Set RsAux = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux(0) = lCant And RsAux(1) = lSum Then
            lCant = 0
        End If
        RsAux.Close
        
    Else
        frm_SetDetalleLogo "Eliminando ..."
        Cons = "Delete * From Marca"
        cAccess.Execute (Cons)
    End If
    
    If lCant > 0 Then
        
        frm_SetDetalleLogo "Pasando datos ..."
        Me.Refresh
        
        pbParcial.Max = lCant + 1       'Le sumo 1 porque borro abajo.
        Cons = "Select * From Marca"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            
            'Veo si la especie existe, sino la inserto.
            Cons = "Select * From Marca Where MarCodigo = " & RsAux("MarCodigo")
            Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If rsAcc.EOF Then
                rsAcc.AddNew
                rsAcc("MarCodigo") = RsAux("MarCodigo")
            Else
                rsAcc.Edit
            End If
            rsAcc("MarNombre") = RsAux("MarNombre")
            rsAcc.Update
            rsAcc.Close
            
            If sMarca <> "" Then sMarca = sMarca & ", "
            sMarca = sMarca & RsAux("MarCodigo")
            
            RsAux.MoveNext
            AumentoProgressParcial
        Loop
        RsAux.Close
        
        frm_SetDetalleLogo "Eliminando ..."
        Cons = "Delete * From Marca Where MarCodigo Not In (" & sMarca & ")"
        cAccess.Execute (Cons)
        
        AumentoProgressParcial
    End If
    
    frm_SetDetalleLogo ""
    pbParcial.Value = 0
    Exit Function
    
errPM:
    acc_ActualizarMarca = f_SetError("Marcas")
    pbParcial.Value = 0
End Function

Private Sub ProcesoPreciosPorArticulo()
On Error GoTo errPPPA
Dim sArt2 As String
Dim cTC As Currency
Dim lPos As Long, idArtAnt As Long

    sArt2 = ""
    'Busco si hay artículos en Access que no tengan precio
    Cons = "Select * From ListaArticulos Where LAArtID Not IN(Select Distinct(PVIArticulo) From PrecioVigente) And LAEsCombo = 0"
    Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAcc.EOF
        If sArt2 <> "" Then sArt2 = sArt2 & ", "
        sArt2 = sArt2 & rsAcc!LAArtID
        rsAcc.MoveNext
    Loop
    rsAcc.Close
    
    ReDim arrArtPre(0)
    If sArt2 <> "" Then
    
        Cons = "Select * From PrecioVigente, TipoCuota, TipoPlan  " _
            & " Where PViArticulo IN(" & sArt2 & ")" _
            & " And PViHabilitado <> 0 And PViMoneda = " & paMonedaPesos _
            & " And PViTipoCuota = TCuCodigo And PViPlan = PlaCodigo" _
            & " Order by PViArticulo"

        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        idArtAnt = 0
        Do While Not RsAux.EOF
            
            'Si no esta en la tabla no lo ingreso.
                
            If idArtAnt <> RsAux!PViArticulo Then
                
                'x las dudas lo agrego a art.modificados.
                If Trim(sArtMod) <> "" Then sArtMod = sArtMod & ", "
                sArtMod = sArtMod & RsAux!PViArticulo
                
                AumentoProgressParcial
                idArtAnt = RsAux!PViArticulo
                
                '1Ero. elimino todos los precios.
                Cons = "Delete * From PrecioVigente Where PViArticulo = " & idArtAnt
                cAccess.Execute (Cons)
                
                Cons = "Select * From ListaArticulos Where LAArtID = " & RsAux!PViArticulo
                Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If rsAcc.EOF Then
                    rsAcc.Close
                    InsertoArticuloPorPrecio RsAux!PViArticulo
                End If
            End If
            
            Cons = "Select * From PrecioVigente Where PViArticulo = " & idArtAnt
            Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            rsAcc.AddNew
            rsAcc("PViArticulo") = RsAux("PViArticulo")
            rsAcc("PViTipoCuota") = RsAux("PViTipoCuota")
            rsAcc("PViCantCuota") = RsAux("TCuCantidad")
            rsAcc("PViDescripcion") = RsAux("TCuAbreviacion")
            'El precio que guardo es el precio de la cuota.
            rsAcc("PViPrecio") = Format(RsAux("PViPrecio") / RsAux("TCuCantidad"), "###0")
            rsAcc.Update
            rsAcc.Close
            If RsAux("TCuVencimientoC") = 0 And RsAux("TCuCantidad") > 0 Then
                InsertoEnArray RsAux("PViArticulo"), RsAux("PViPrecio") / RsAux("TCuCantidad"), RsAux("PlaNombre"), Trim(RsAux("TCuAbreviacion")), RsAux("TCuCodigo")
            End If
    
            RsAux.MoveNext
        Loop
        RsAux.Close
        
    
        'Actualizo el precio en la tabla ListaArticulos
        For lPos = 1 To UBound(arrArtPre)
            Cons = "Select * From ListaArticulos Where LAArtID = " & arrArtPre(lPos).idArt
            Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rsAcc.EOF Then
                rsAcc.Edit
                rsAcc("LAContado") = arrArtPre(lPos).Ctdo
                rsAcc("LAPlan") = Trim(arrArtPre(lPos).Plan)
                If arrArtPre(lPos).CuotaFinanciado = -1 Then
                    rsAcc("LAFinanciado") = Null
                    rsAcc("LATextoFinanciado") = Null
                Else
                    rsAcc("LAFinanciado") = arrArtPre(lPos).CuotaFinanciado
                    rsAcc("LATextoFinanciado") = arrArtPre(lPos).TCuotaAbrev
                End If
                rsAcc.Update
            End If
            rsAcc.Close
            AumentoProgressParcial
        Next lPos
    End If
    
errPPPA:
End Sub

Private Function PasoPrecioVigente() As String
On Error GoTo errPV
Dim cTC As Currency
Dim lPos As Long, idArtAnt As Long
    
    PasoPrecioVigente = ""
        
    If opHacer(0).Value Then
        'Corro todo
        lLogoPaso.Caption = "Precios Vigentes"
        frm_SetDetalleLogo "Validando ..."
        acc_PrecioVigenteTotal
        Exit Function
        
    Else
        'No paso los precios x más que hayan artículos modificados.
        If chPrecio.Value = 0 And opHacer(1).Value Then Exit Function
    End If
    
    lLogoPaso.Caption = "Precios Vigentes"
    frm_SetDetalleLogo "Validando ..."

    Dim dFAct As Date
    
    dFAct = DateAdd("n", -5, CDate(paFUltActualizacion))
        
    'Todos los modificados y aquellos que los precios se modificaron.
    Cons = "Select Count(Distinct(PVIArticulo)) From PrecioVigente " _
        & " Where PViVigencia <= '" & Format(Now, "mm/dd/yyyy hh:nn:ss") & "'" _
        & " And PViHabilitado <> 0  And PViMoneda = " & paMonedaPesos _
        & " And (PViArticulo in (" & sArtMod & ") or PViVigencia >= '" & Format(dFAct, "mm/dd/yyyy hh:nn:ss") & "')"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    lPos = RsAux(0)
    RsAux.Close
    
    If lPos = 0 Then
        ProcesoPreciosPorArticulo
        pbParcial.Value = 0
        Exit Function
    Else
        pbParcial.Max = (lPos * 2) + 1
    End If
    
    ReDim arrArtPre(0)
    frm_SetDetalleLogo "Pasando ..."
    
    Cons = "Select * From PrecioVigente, TipoCuota, TipoPlan  " _
        & " Where PViVigencia < '" & Format(Now, "mm/dd/yyyy hh:nn:ss") & "'" _
        & " And PViHabilitado <> 0  And PViMoneda = " & paMonedaPesos _
        & " And (PViArticulo in (" & sArtMod & ") or PViVigencia >= '" & Format(dFAct, "mm/dd/yyyy hh:nn:ss") & "')" _
        & " And PViTipoCuota = TCuCodigo And PViPlan = PlaCodigo" _
        & " Order by PViArticulo"
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    idArtAnt = 0
    Do While Not RsAux.EOF
        
        'Si no está en la tabla no lo ingreso.
            
        If idArtAnt <> RsAux!PViArticulo Then
            
            'x las dudas lo agrego a art.modificados.
            If Trim(sArtMod) <> "" Then sArtMod = sArtMod & ", "
            sArtMod = sArtMod & RsAux!PViArticulo
            
            AumentoProgressParcial
            idArtAnt = RsAux!PViArticulo
            '1Ero. elimino todos los precios.
            Cons = "Delete * From PrecioVigente Where PViArticulo = " & idArtAnt
            cAccess.Execute (Cons)
            
            Cons = "Select * From ListaArticulos Where LAArtID = " & RsAux!PViArticulo
            Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If rsAcc.EOF Then
                rsAcc.Close
                InsertoArticuloPorPrecio RsAux!PViArticulo
            End If
        End If
        
        Cons = "Select * From PrecioVigente Where PViArticulo = " & idArtAnt
        Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        rsAcc.AddNew
        rsAcc("PViArticulo") = RsAux("PViArticulo")
        rsAcc("PViTipoCuota") = RsAux("PViTipoCuota")
        rsAcc("PViCantCuota") = RsAux("TCuCantidad")
        rsAcc("PViDescripcion") = RsAux("TCuAbreviacion")
        'El precio que guardo es el precio de la cuota.
        rsAcc("PViPrecio") = Format(RsAux("PViPrecio") / RsAux("TCuCantidad"), "###0")
        rsAcc.Update
        rsAcc.Close
        If RsAux("TCuVencimientoC") = 0 And RsAux("TCuCantidad") > 0 Then
            InsertoEnArray RsAux("PViArticulo"), RsAux("PViPrecio") / RsAux("TCuCantidad"), RsAux("PlaNombre"), Trim(RsAux("TCuAbreviacion")), RsAux("TCuCodigo")
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    

    'Actualizo el precio en la tabla ListaArticulos
    For lPos = 1 To UBound(arrArtPre)
        Cons = "Select * From ListaArticulos Where LAArtID = " & arrArtPre(lPos).idArt
        Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsAcc.EOF Then
            rsAcc.Edit
            rsAcc("LAContado") = arrArtPre(lPos).Ctdo
            rsAcc("LAPlan") = Trim(arrArtPre(lPos).Plan)
            If arrArtPre(lPos).CuotaFinanciado = -1 Then
                rsAcc("LAFinanciado") = Null
                rsAcc("LATextoFinanciado") = Null
            Else
                rsAcc("LAFinanciado") = arrArtPre(lPos).CuotaFinanciado
                rsAcc("LATextoFinanciado") = arrArtPre(lPos).TCuotaAbrev
            End If
            rsAcc.Update
        End If
        rsAcc.Close
        AumentoProgressParcial
    Next lPos
    
    ProcesoPreciosPorArticulo
    
    pbParcial.Value = pbParcial.Max
    pbParcial.Value = 0
    Exit Function
    
errPV:
    PasoPrecioVigente = "(PasoPrecioVigente) " & Trim(Err.Description)
    pbParcial.Value = 0
End Function

Private Sub acc_PrecioVigenteTotal()
Dim cTC As Currency
Dim lPos As Long, idArtAnt As Long

    Cons = "Delete * From PrecioVigente"
    cAccess.Execute (Cons)
    
    Cons = "Select Count(Distinct(PVIArticulo)) From PrecioVigente " _
        & " Where PViHabilitado <> 0 And PViMoneda = " & paMonedaPesos _
        & " And PVIArticulo IN (Select ArtID From Articulo Where ArtEnUso = 1)"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    lPos = RsAux(0)
    RsAux.Close
    
    If lPos = 0 Then
        pbParcial.Value = 0
        Exit Sub
    Else
        pbParcial.Max = (lPos * 2) + 1
    End If
        
    ReDim arrArtPre(0)
    frm_SetDetalleLogo "Pasando ..."
    
    'Saco los art. cuyo precio vigente sea mayor = a la fecha de act. y menor a ahora.
    Cons = "Select * From PrecioVigente, TipoCuota, TipoPlan  " _
        & " Where PViHabilitado <> 0 And PViMoneda = " & paMonedaPesos _
        & " And PViTipoCuota = TCuCodigo And PViPlan = PlaCodigo" _
        & " And PVIArticulo IN (Select ArtID From Articulo Where ArtEnUso = 1)" _
        & " Order by PViArticulo"
    
    cBase.QueryTimeout = 60
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    cBase.QueryTimeout = 20
    idArtAnt = 0
    Do While Not RsAux.EOF
        
        'Si no esta en la tabla no lo ingreso.
        If idArtAnt <> RsAux!PViArticulo Then
            AumentoProgressParcial
            idArtAnt = RsAux!PViArticulo
            
            Cons = "Select * From ListaArticulos Where LAArtID = " & RsAux!PViArticulo
            Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If rsAcc.EOF Then
                rsAcc.Close
                InsertoArticuloPorPrecio RsAux!PViArticulo
            End If
        End If
        
        Cons = "Select * From PrecioVigente Where PViArticulo = " & idArtAnt
        Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        rsAcc.AddNew
        rsAcc("PViArticulo") = RsAux("PViArticulo")
        rsAcc("PViTipoCuota") = RsAux("PViTipoCuota")
        rsAcc("PViCantCuota") = RsAux("TCuCantidad")
        rsAcc("PViDescripcion") = RsAux("TCuAbreviacion")
        'El precio que guardo es el precio de la cuota.
        rsAcc("PViPrecio") = Format(RsAux("PViPrecio") / RsAux("TCuCantidad"), "###0")
        rsAcc.Update
        rsAcc.Close
        If RsAux("TCuVencimientoC") = 0 And RsAux("TCuCantidad") > 0 Then
            InsertoEnArray RsAux("PViArticulo"), RsAux("PViPrecio") / RsAux("TCuCantidad"), RsAux("PlaNombre"), Trim(RsAux("TCuAbreviacion")), RsAux("TCuCodigo")
        End If

        RsAux.MoveNext
    Loop
    RsAux.Close
    

    'Actualizo el precio en la tabla ListaArticulos
    For lPos = 1 To UBound(arrArtPre)
        Cons = "Select * From ListaArticulos Where LAArtID = " & arrArtPre(lPos).idArt
        Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsAcc.EOF Then
            rsAcc.Edit
            rsAcc("LAContado") = arrArtPre(lPos).Ctdo
            rsAcc("LAPlan") = Trim(arrArtPre(lPos).Plan)
            If arrArtPre(lPos).CuotaFinanciado = -1 Then
                rsAcc("LAFinanciado") = Null
                rsAcc("LATextoFinanciado") = Null
            Else
                rsAcc("LAFinanciado") = arrArtPre(lPos).CuotaFinanciado
                rsAcc("LATextoFinanciado") = arrArtPre(lPos).TCuotaAbrev
            End If
            rsAcc.Update
        End If
        rsAcc.Close
        AumentoProgressParcial
    Next lPos
    pbParcial.Value = pbParcial.Max
    pbParcial.Value = 0
    

End Sub

Private Sub InsertoEnArray(ByVal lArt As Long, ByVal cPrecio As Currency, _
                        ByVal sPlan As String, sAbrevTC As String, lCodigoTC As Long)
Dim lPos As Long, lCont As Long

    lPos = -1
        
    For lCont = 1 To UBound(arrArtPre)
        If arrArtPre(lCont).idArt = lArt Then
            lPos = lCont
            Exit For
        End If
    Next

    If lPos = -1 Then
        lPos = UBound(arrArtPre) + 1
        ReDim Preserve arrArtPre(lPos)
        arrArtPre(lPos).idArt = lArt
        arrArtPre(lPos).TCuotaAbrev = ""
        arrArtPre(lPos).CuotaFinanciado = -1
    End If
    
    arrArtPre(lPos).Plan = sPlan
    If lCodigoTC = paCuotaCtdo Then
        arrArtPre(lPos).Ctdo = cPrecio
    Else
        'arrartpre(lpos).CuotaFinanciado = -1
        If (arrArtPre(lPos).CuotaFinanciado > cPrecio Or arrArtPre(lPos).CuotaFinanciado = -1) _
                                    And cPrecio > paCuotaMin Then
            arrArtPre(lPos).CuotaFinanciado = Format(cPrecio, "###0")
            arrArtPre(lPos).TCuotaAbrev = sAbrevTC
        End If
    End If
    
End Sub

Private Function acc_ActualizarGlosario() As String
On Error GoTo errPC
Dim sTexto As String
Dim lSum As Long
    
    acc_ActualizarGlosario = ""
    If Not (chGlosario.Value = 1 Or opHacer(0).Value) Then Exit Function
    
    lLogoPaso.Caption = "Glosario."
    lLogoDet.Caption = "Validando ..."
    
    Cons = "Select Count(*)  From Glosario"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux(0) = 0 Then
        RsAux.Close
        Cons = "Delete * From Glosario"
        cAccess.Execute (Cons)
        Exit Function
    Else
        pbParcial.Max = RsAux(0) + 2
        RsAux.Close
    End If
    
    
    pbParcial.Value = 1
    Cons = "Delete * From Glosario"
    cAccess.Execute (Cons)
    pbParcial.Value = 2
    
    lLogoDet.Caption = "Pasando datos ..."
    
    Cons = "Select * From Glosario"
    Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
    Cons = "Select * From Glosario"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        rsAcc.AddNew
        rsAcc("GloID") = RsAux("GloID")
        rsAcc("GloNombre") = RsAux("GloNombre")
        If Not IsNull(RsAux("GloAlto")) Then
            rsAcc("GloAlto") = RsAux("GloAlto")
        End If
        If Not IsNull(RsAux("GloAncho")) Then
            rsAcc("GloAncho") = RsAux("GloAncho")
        End If
        sTexto = RetornoLongText("GloTexto")
        If sTexto <> "" Then rsAcc("GloTexto") = sTexto
        rsAcc("GloScroll") = RsAux("GloScroll")
        rsAcc.Update
        
        RsAux.MoveNext
        AumentoProgressParcial
    Loop
    RsAux.Close
    pbParcial.Value = 0
    Exit Function
    
errPC:
    acc_ActualizarGlosario = f_SetError("Glosario")
End Function

Private Function acc_ActualizoStock() As String
On Error GoTo errAS
Dim lSanos As Long, lVta As Long, lCont As Long
Dim iStock As Integer
Dim boEsCombo As Boolean
Dim sFEmb As String
Dim sArt As String, ArrArt() As tArticulo

    acc_ActualizoStock = ""
    
    If Not (chStock.Value = 1 Or sArtNew <> "" Or opHacer(0).Value Or _
        Abs(DateDiff("n", Now, paFUltActualizacion)) > 60 Or Val(tArticulo.Tag) > 0) Then
        Exit Function
    End If
    
    lLogoPaso.Caption = "Stock"
    frm_SetDetalleLogo "Validando ..."
    
    If Val(tArticulo.Tag) = 0 Then
        'No elimino los de servicios ya que tengo que ponerle 1.
        Cons = "Select Count(*) From ListaArticulos Where LATipo <> " & paTipoServ _
            & " Or (LATipo = " & paTipoServ & " And LAHayStock <> 1)" _
            & " Or (LATipo = " & paTipoServ & " And LAFechaArribo Is Not Null)"
            
        Set RsAux = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux(0) = 0 Then
            RsAux.Close
            Exit Function
        Else
            pbParcial.Max = RsAux(0)
            RsAux.Close
        End If
    Else
        pbParcial.Max = 1
    End If
    
    
    If Val(tArticulo.Tag) = 0 Then
    
        'Para todos los artículos recorro la tabla de stock y calculo la vtas de los últimos 5 días.
        Cons = "Select * From ListaArticulos Where LATipo <> " & paTipoServ _
            & " Or (LATipo = " & paTipoServ & " And LAHayStock <> 1)" _
            & " Or (LATipo = " & paTipoServ & " And LAFechaArribo Is Not Null)"
            
        Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        rdoErrors.Clear
        
        If rsAcc.EOF Then
            rsAcc.Close
            Exit Function
        Else
            frm_SetDetalleLogo "Pasando ..."
            Do While Not rsAcc.EOF
                If Not rsAcc("LAEsCombo") And rsAcc("LATipo") <> paTipoServ And rsAcc("LAHabilitado") Then
                    If sArt = "" Then
                        sArt = rsAcc("LAArtID")
                    Else
                        sArt = sArt & ", " & rsAcc("LAArtID")
                    End If
                End If
                rsAcc.MoveNext
            Loop
            rsAcc.MoveFirst
        End If
    Else
        'No se si es del tipo servicio.
        Cons = "Select * From ListaArticulos Where LATipo <> " & paTipoServ & " And LAArtID = " & Val(tArticulo.Tag)
        Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        rdoErrors.Clear
        If rsAcc.EOF Then
            rsAcc.Close
            Exit Function
        End If
        sArt = Val(tArticulo.Tag)
    End If
    
    ReDim ArrArt(0)
    
    Cons = "Select ArtID, Sum(AArCantidadNCo) + Sum(AArCantidadNCr) + Sum(AArCantidadECo) + Sum(AArCantidadECr) as Vta, StTCantidad " _
        & " From Articulo Left Outer Join AcumuladoArticulo ON ArtID = AArArticulo And AArFecha >= '" & Format(DateAdd("d", paStockParaXDias * -1, Date), "mm/dd/yyyy 00:00:00") & "'" _
        & " Left Outer Join StockTotal On ArtID = StTArticulo And StTEstado = " & paEstSano _
        & " Where ArtID In (" & sArt & ") Group By ArtID, StTcantidad Order by ArtID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        ReDim Preserve ArrArt(UBound(ArrArt) + 1)
        With ArrArt(UBound(ArrArt))
            .ID = RsAux(0)
            If Not IsNull(RsAux(1)) Then .Vta = RsAux(1) Else .Vta = 0
            If Not IsNull(RsAux(2)) Then .Sanos = RsAux(2) Else .Sanos = 0
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Do While Not rsAcc.EOF
        
        If rsAcc("LATipo") <> paTipoServ And rsAcc("LAHabilitado") Then
            lSanos = 0: lVta = 0
            boEsCombo = rsAcc("LAEsCombo")
            If boEsCombo Then
                lSanos = StockSanoEnCombo(rsAcc("LAArtID"))
            Else
                'Busco en el array el articulo.
                For lCont = 1 To UBound(ArrArt)
                    If ArrArt(lCont).ID = rsAcc("LAArtID") Then
                        lSanos = ArrArt(lCont).Sanos
                        lVta = ArrArt(lCont).Vta
                        Exit For
                    End If
                Next
            End If
            
            If lSanos > 5 Then
                If boEsCombo Then
                    lVta = VentaUltimosDiasArticuloCombo(rsAcc("LAArtID"))
                End If
                
                If lSanos < lVta Then
                    iStock = 2
                Else
                    iStock = 1
                End If
                
            Else
                
                If lSanos = 0 Then
                    iStock = 0
                Else
                    iStock = 2
                End If
                
            End If
            
        Else
            
            If rsAcc("LATipo") = paTipoServ Then
                iStock = 4      'Carlos :10-4-02
            Else
                iStock = 0
            End If
            
        End If
        rsAcc.Edit
        
        If iStock <> rsAcc("LAHayStock") And (paWebStock Or paIntraStock) Then
            If sArtMod = "" Then sArtMod = rsAcc("LAArtID") Else sArtMod = sArtMod & ", " & rsAcc("LAArtID")
        End If
        rsAcc("LAHayStock") = iStock
        
        If iStock = 1 Then
            rsAcc("LAFechaArribo") = Null
        Else
            sFEmb = FechaDeEmbarque(rsAcc("LAArtID"), boEsCombo)
            If IsDate(sFEmb) Then
                rsAcc("LAFechaArribo") = Format(sFEmb, "mm/dd/yyyy")
            Else
                rsAcc("LAFechaArribo") = Null
            End If
        End If
        rsAcc.Update
        
        rsAcc.MoveNext
        If pbParcial.Max > pbParcial.Value + 2 Then AumentoProgressParcial
    Loop
    rsAcc.Close
    Erase ArrArt
    
    pbParcial.Value = 0
    Exit Function
    
errAS:
    acc_ActualizoStock = f_SetError("Stock")
    pbParcial.Value = 0
End Function
Private Function EsArtCombo(ByVal lArt As Long) As Boolean
On Error GoTo errAC

    EsArtCombo = False
    Cons = "Select * From Articulo Where ArtID = " & lArt
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If RsAux("ArtEsCombo") Then EsArtCombo = True
    End If
    RsAux.Close
    Me.Refresh
errAC:

End Function
Private Function StockSanoEnCombo(ByVal lArt As Long) As Long
On Error GoTo errSS
    StockSanoEnCombo = 0
    Cons = "Select Min(StTcantidad / ParCantidad) From Presupuesto, PresupuestoArticulo, StockTotal" _
        & " Where PReArtCombo = " & lArt & " And sttEstado = " & paEstSano _
        & " And PReID = PArPresupuesto And PArArticulo = StTArticulo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If RsAux(0) > 0 Then StockSanoEnCombo = RsAux(0)
    End If
    RsAux.Close
errSS:
End Function

Private Function VentaUltimosDiasArticuloCombo(ByVal lArt As Long) As Long
On Error GoTo errVUAC

'Retorno el artículo que tuvo + vtas de todos los del combo.
    VentaUltimosDiasArticuloCombo = 0
    
    Cons = "Select Sum(RenCantidad), RenArticulo " _
        & " From Renglon, Documento, Presupuesto, PresupuestoArticulo " _
         & " Where DocFecha >= '" & Format(DateAdd("d", paStockParaXDias * -1, Date), "mm/dd/yyyy 00:00:00") & "'" _
         & " And DocTipo IN (1,2) And DocAnulado = 0 And PreArtCombo = " & lArt _
         & " And DocCodigo = RenDocumento And RenArticulo = PArArticulo And PArPresupuesto = PreID" _
         & " Group By RenArticulo Order by 1 Desc"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        If Not IsNull(RsAux(0)) Then VentaUltimosDiasArticuloCombo = RsAux(0)
    End If
    RsAux.Close

errVUAC:
End Function

Private Function FechaDeEmbarque(ByVal lArt As Long, ByVal boEsCombo As Boolean) As String
On Error GoTo errFE
Dim lDif As Long, lResto As Long
    FechaDeEmbarque = ""
    If boEsCombo Then
        Cons = "Select Min(EmbFAPrometido) " _
            & " From ArticuloFolder, Embarque, Presupuesto, PresupuestoArticulo " _
            & " Where PreArtCombo = " & lArt & " And AFoTipo = 2 And EmbFAPrometido Is Not Null " _
            & " And EmbFArribo Is Null And AFoCodigo = EmbID And AFoArticulo = PArArticulo And PreID = PArPresupuesto "
    Else
        Cons = "Select Min(EmbFAPrometido) From ArticuloFolder, Embarque " _
            & " Where AFoArticulo = " & lArt & " And AFoTipo = 2 And EmbFAPrometido Is Not Null " _
            & " And AFoCodigo = EmbID And EmbFArribo Is Null"
    End If
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux(0)) Then
            'Utilizo porcentaje de embarque si hay.
            If paPorcFEmb = 0 Then
                FechaDeEmbarque = RsAux(0)
            Else
                If RsAux(0) <= Date Then
                    FechaDeEmbarque = RsAux(0)
                Else
                    lDif = Abs(DateDiff("d", RsAux(0), Date))
                    If Abs(lDif * paPorcFEmb - CInt(lDif * paPorcFEmb)) < 0.5 Then
                        lDif = CInt(lDif * paPorcFEmb) + 1
                    Else
                        lDif = CInt(lDif * paPorcFEmb)
                    End If
                    FechaDeEmbarque = DateAdd("d", lDif, Date)
                End If
            End If
        End If
    End If
    RsAux.Close
    
errFE:
End Function

Private Function db_SetUltimaEjecucion() As String
On Error GoTo errGUE
    
    db_SetUltimaEjecucion = ""
    lLogoPaso.Caption = "Última Ejecución ..."
    frm_SetDetalleLogo "Grabando ..."
    
    pbParcial.Max = 2
    
    Cons = "Select * From Parametro Where ParNombre = 'WebFechaActualizada'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    rdoErrors.Clear
    
    pbParcial.Value = 1
    If RsAux.EOF Then
        RsAux.AddNew
        RsAux("ParNombre") = "WebFechaActualizada"
    Else
        RsAux.Edit
    End If
    FechaDelServidor
    paFUltActualizacion = Format(Now, "dd/mm/yyyy hh:nn:ss")
    RsAux("ParTexto") = paFUltActualizacion
    RsAux.Update
    RsAux.Close
    pbParcial.Value = 2
    pbParcial.Value = 0
    Exit Function
errGUE:
    db_SetUltimaEjecucion = f_SetError("Última ejecución")
End Function

Private Function acc_ActualizarListasDePrecios() As String
On Error GoTo errPLP
Dim sError As String, sPlaTexto As String
    
    acc_ActualizarListasDePrecios = ""
    If chLista.Value = 0 And opHacer(1).Value Then Exit Function
    
    lLogoPaso.Caption = "Listas de Precios"
    frm_SetDetalleLogo "Validando ..."
    rdoErrors.Clear
    Cons = "Select Count(*) From CodigoTexto, logdb.dbo.Plantilla Where Tipo = 68 And Puntaje = 1" _
        & " And Clase = PlaCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    rdoErrors.Clear
    If RsAux(0) = 0 Then
        RsAux.Close
        Exit Function
    Else
        pbParcial.Max = RsAux(0)
        RsAux.Close
    End If
    
    Cons = "Select * From CodigoTexto, logdb.dbo.Plantilla Where Tipo = 68 And Puntaje = 1" _
        & " And Clase = PlaCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    rdoErrors.Clear
    Do While Not RsAux.EOF
        sError = GeneroPlantillaListaPrecio
        If sError <> "" Then
            acc_ActualizarListasDePrecios = acc_ActualizarListasDePrecios & sError
        End If
        RsAux.MoveNext
        AumentoProgressParcial
    Loop
    RsAux.Close
    
    If acc_ActualizarListasDePrecios <> "" Then
        acc_ActualizarListasDePrecios = String(30, ".") & "Listas de Precios" & vbCrLf & acc_ActualizarListasDePrecios & String(30, ".") & "Listas de Precios" & vbCrLf
    End If
    Exit Function
    
errPLP:
    acc_ActualizarListasDePrecios = f_SetError("Listas de Precios")
End Function

Private Function GeneroPlantillaListaPrecio() As String
On Error GoTo errGP
Dim sTexto As String, sAsunto As String, sFile As String, sErr As String
Dim sParam As String

    Dim objPla As New clsPlantillaI
    sTexto = ""
    GeneroPlantillaListaPrecio = ""
    sParam = ""
    If Not IsNull(RsAux!Valor1) Then sParam = RsAux!Valor1
    
    If objPla.ProcesoPlantillaInteractiva(cBase, RsAux("PlaCodigo"), 0, sTexto, sAsunto, sParam, False) Then
        
        If sAsunto = "" Then
            GeneroPlantillaListaPrecio = "Plantilla PlaCodigo = " & RsAux("PlaCodigo") & " NO TIENE ASUNTO" & vbCrLf
        Else
            If Right(RsAux("Texto2"), 1) <> "\" Then
                sFile = Trim(RsAux("Texto2")) & "\" & sAsunto
            Else
                sFile = Trim(RsAux("Texto2")) & sAsunto
            End If
            
            sErr = GraboArchivo(sFile, sTexto)
            If sErr <> "" Then
                If sErr = "75" Then
                    'Espero 1 segundo y reintento.
                    Sleep 1000
                    sErr = GraboArchivo(sFile, sTexto)
                    If sErr <> "" Then
                        GeneroPlantillaListaPrecio = "Plantilla PlaCodigo = " & RsAux("PlaCodigo") & " Error al grabar el archivo: " & sErr & vbCrLf
                    End If
                Else
                    GeneroPlantillaListaPrecio = "Plantilla PlaCodigo = " & RsAux("PlaCodigo") & " Error al grabar el archivo " & vbCrLf
                End If
            End If
            
        End If
        
    Else
        GeneroPlantillaListaPrecio = "Plantilla  PlaCodigo = " & RsAux("PlaCodigo") & "Error al generar la plantilla." & vbCrLf
    End If
    GoTo evFin
    
errGP:
    GeneroPlantillaListaPrecio = "Plantilla  PlaCodigo = " & RsAux("PlaCodigo") & " " & Trim(Err.Description) & vbCrLf
    
evFin:
    Set objPla = Nothing
    Exit Function
End Function

Private Function GeneroPaginasDeArticulos(ByVal sUltAct As String) As String
On Error GoTo errGPA
Dim lPos As Long, lCont As Long
Dim sErr As String

    GeneroPaginasDeArticulos = ""
    lLogoPaso.Caption = "Páginas Artículos"
    frm_SetDetalleLogo "Generando ..."

    ReDim arrArtPla(0)
    
    '1ero Busco las Plantillas Modificadas o los artículos modificados.
    '2do con el array genero los archivos en base a las plantillas.
    Cons = "Select AWPArticulo, AWPPlantillaIntra, AWPPlantillaWeb From Plantilla, ArticuloWebPage " _
            & "Where PlaModificada >= '" & Format(CDate(sUltAct), "mm/dd/yyyy hh:nn:ss") & "'" _
            & " And (PlaCodigo = AWPPlantillaWeb Or PlaCodigo = AWPPlantillaIntra) "
    
    If Trim(sArtMod) <> "" And sArtMod <> "0" Then
        Cons = Cons _
        & "Union " _
        & "Select AWPArticulo, AWPPlantillaIntra, AWPPlantillaWeb From Plantilla, ArticuloWebPage " _
        & " Where AWPArticulo IN (" & sArtMod & ") And (PlaCodigo = AWPPlantillaWeb Or PlaCodigo = AWPPlantillaIntra)"
    End If
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        lPos = EstaEnArray(RsAux("AWPArticulo"))
        If lPos = 0 Then
            lPos = UBound(arrArtPla) + 1
            ReDim Preserve arrArtPla(lPos)
        End If
        arrArtPla(lPos).idArticulo = RsAux("AWPArticulo")
        If Not IsNull(RsAux("AWPPlantillaWeb")) Then arrArtPla(lPos).idPlaWeb = RsAux("AWPPlantillaWeb")
        If Not IsNull(RsAux("AWPPlantillaIntra")) Then arrArtPla(lPos).idPlaIntra = RsAux("AWPPlantillaIntra")
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If UBound(arrArtPla) > 0 Then
        
        'Busco en los archivos aquellos que contengan el que voy a generar.
        AgregoArticuloDeArchivo
        
        pbParcial.Max = UBound(arrArtPla)
        For lCont = 1 To UBound(arrArtPla)
            If arrArtPla(lCont).idPlaIntra > 0 Then
                sErr = GeneroPlantillaArticulos(paPathIntra, arrArtPla(lCont).idPlaIntra, arrArtPla(lCont).idArticulo)
                If sErr <> "" Then GeneroPaginasDeArticulos = GeneroPaginasDeArticulos & vbCrLf & sErr
            End If
            If arrArtPla(lCont).idPlaWeb > 0 Then
                sErr = GeneroPlantillaArticulos(paPathWeb, arrArtPla(lCont).idPlaWeb, arrArtPla(lCont).idArticulo)
                If sErr <> "" Then GeneroPaginasDeArticulos = GeneroPaginasDeArticulos & vbCrLf & sErr
            End If
            pbParcial.Value = lCont
        Next
        pbParcial.Value = 0
    End If
    If GeneroPaginasDeArticulos <> "" Then
        GeneroPaginasDeArticulos = String(30, ".") & "Páginas htm" & vbCrLf & GeneroPaginasDeArticulos & String(30, ".") & "Páginas htm" & vbCrLf
    End If
    Exit Function
    
errGPA:
   GeneroPaginasDeArticulos = f_SetError("Genero Páginas")
End Function

Private Function EstaEnArray(ByVal lArt As Long) As Long
Dim lCont As Long
    
    EstaEnArray = 0
    For lCont = 1 To UBound(arrArtPla)
        If arrArtPla(lCont).idArticulo = lArt Then
            EstaEnArray = lCont
            Exit Function
        End If
    Next
    
End Function

Private Function GeneroPlantillaArticulos(ByVal sPath As String, ByVal lPlantilla As Long, ByVal sArt As String) As String
On Error GoTo errGP
Dim sTexto As String, sAsunto As String, sErr As String

    Dim objPla As New clsPlantillaI
    sTexto = ""
    GeneroPlantillaArticulos = ""
    sPath = Replace(sPath, "[id]", Val(sArt), , , vbTextCompare)
    
    If objPla.ProcesoPlantillaInteractiva(cBase, lPlantilla, 0, sTexto, sAsunto, sArt, False) Then
        
        sErr = GraboArchivo(sPath, sTexto)
        If sErr <> "" Then
            If sErr = "75" Then
                'Espero 1/2 segundo y reintento.
                Sleep 500
                sErr = GraboArchivo(sPath, sTexto)
                If sErr <> "" Then
                    GeneroPlantillaArticulos = "Plantilla PlaCodigo = " & lPlantilla & "IDArticulo = " & Trim(sArt) & " Error al grabar el archivo " & vbCrLf
                End If
            Else
                GeneroPlantillaArticulos = "Plantilla PlaCodigo = " & lPlantilla & "IDArticulo = " & Trim(sArt) & " Error al grabar el archivo " & vbCrLf
            End If
        End If
    Else
        GeneroPlantillaArticulos = "Plantilla PlaCodigo = " & lPlantilla & "IDArticulo = " & Trim(sArt) & " ERROR al generar la plantilla." & vbCrLf
    End If
    GoTo evFin
    
errGP:
    GeneroPlantillaArticulos = "Plantilla PlaCodigo = " & lPlantilla & "IDArticulo = " & Trim(sArt) & Trim(Err.Description) & vbCrLf
    
evFin:
    Set objPla = Nothing
    Exit Function
    
End Function

Private Function GraboArchivo(ByVal sFile As String, sTexto As String) As String
On Error GoTo errGA
    Open sFile For Output As #1
    Print #1, sTexto
    Close #1
    Exit Function
errGA:
    If Err.Number = 75 Then GraboArchivo = "75" Else GraboArchivo = Err.Description
End Function

Private Sub GraboErrores()
On Error Resume Next
Dim lCont As Long
Dim sError As String
    
    For lCont = 1 To colError.Count
        sError = colError.Item(lCont) & vbCrLf
    Next lCont
    sError = "Ejecución: " & Format(Now, "Ddd, d/mm/yy hh:nn:ss") & vbCrLf & sError
    GraboArchivo sFileErr, sError
    
End Sub
Private Sub CargoArchivosWebIntra(ByRef arrWebIntra() As tArtHtmAsp, ByVal sPathWI As String)
Dim sArch As String
Dim sExt As String, sPath As String, sClave As String, sPathTotal As String

    Screen.MousePointer = 11
    
    sExt = Mid(sPathWI, InStrRev(sPathWI, ".") + 1)
    sPath = Mid(sPathWI, 1, InStrRev(sPathWI, "[id].", , vbTextCompare) - 1)
    sPathTotal = Mid(sPathWI, 1, InStrRev(sPathWI, "\", , vbTextCompare))
    sClave = Mid(sPath, InStrRev(sPath, "\", , vbTextCompare) + 1)
    
    sArch = Dir(sPath & "*." & sExt, vbArchive)
    Do While sArch <> ""
        
        'válido que siguiendo la clave venga un número.
        If IsNumeric(Mid(sArch, Len(sClave) + 1, 1)) Then

            ReDim Preserve arrWebIntra(UBound(arrWebIntra) + 1)
            With arrWebIntra(UBound(arrWebIntra))
                .sPath = sArch
                .sData = RetornoStringDeArchivo(sPathTotal & sArch)
            End With
        
        End If
        
        sArch = Dir     'Pido el siguiente
    Loop
    Screen.MousePointer = 0
    
End Sub

Private Function RetornoStringDeArchivo(ByVal sPath As String) As String
On Error GoTo errRA
Dim lFile As Long
    RetornoStringDeArchivo = ""
    lFile = FreeFile
    Open sPath For Input As lFile
    RetornoStringDeArchivo = Input(LOF(lFile), lFile)
    Close lFile
    Exit Function
errRA:
End Function

Private Sub AgregoArticuloDeArchivo()
Dim lCont As Long, lCont1 As Long, lPos As Long
Dim sClave As String, sDescarto As String
Dim sClaveWeb As String, sClaveIntra As String

    ReDim arrFileWeb(0)
    If paWebRelacion Then CargoArchivosWebIntra arrFileWeb, paPathWeb
    
    ReDim arrFileIntra(0)
    If paIntraRelacion Then CargoArchivosWebIntra arrFileIntra, paPathIntra
    
    sClaveWeb = Mid(paPathWeb, InStrRev(paPathWeb, "\", , vbTextCompare) + 1)
    sClaveIntra = Mid(paPathIntra, InStrRev(paPathIntra, "\", , vbTextCompare) + 1)
    
    lCont = 1
    Do While lCont <= UBound(arrArtPla)
        
        If paWebRelacion Then
            sClave = Replace(sClaveWeb, "[id]", arrArtPla(lCont).idArticulo, , , vbTextCompare)
            sDescarto = Mid(sClaveWeb, 1, InStrRev(sClaveWeb, "[id].", , vbTextCompare) - 1)
            
            For lCont1 = 1 To UBound(arrFileWeb)
            
                If InStr(1, arrFileWeb(lCont1).sData, sClave, vbTextCompare) > 0 Then
                    
                    'Agrego si no esta
                    sDescarto = Mid(arrFileWeb(lCont1).sPath, Len(sDescarto) + 1)
                    sDescarto = Mid(sDescarto, 1, InStr(1, sDescarto, ".") - 1)
                    
                    lPos = EstaEnArray(Val(sDescarto))
                    If lPos = 0 And Val(sDescarto) > 0 Then
                        'Tengo que agregarlo.
                        CargoNuevoArchivoArray Val(sDescarto)
                    End If
                End If
            
            Next
        End If
        
        'Intranet
        If paIntraRelacion Then
            sClave = Replace(sClaveIntra, "[id]", arrArtPla(lCont).idArticulo, , , vbTextCompare)
            sDescarto = Mid(sClaveIntra, 1, InStrRev(sClaveIntra, "[id].", , vbTextCompare) - 1)
    
            For lCont1 = 1 To UBound(arrFileIntra)
    
                If InStr(1, arrFileIntra(lCont1).sData, sClave, vbTextCompare) > 0 Then
                    'Al nombre del archivo le saco el id.
                    'Agrego si no esta
                    sDescarto = Mid(arrFileIntra(lCont1).sPath, Len(sDescarto) + 1)
                    sDescarto = Mid(sDescarto, 1, InStr(1, sDescarto, ".") - 1)
                                        
                    'Agrego si no esta
                    lPos = EstaEnArray(Val(sDescarto))
                    If lPos = 0 And Val(sDescarto) > 0 Then
                        'Tengo que agregarlo.
                        CargoNuevoArchivoArray Val(sDescarto)
                    End If
                                    
                End If
            
            Next
        End If
        lCont = lCont + 1
    Loop
    
End Sub

Private Sub CargoNuevoArchivoArray(ByVal idArt As Long)
    
    Cons = "Select AWPArticulo, isnull(AWPPlantillaIntra, 0), isnull(AWPPlantillaWeb, 0) From ArticuloWebPage " _
        & " Where AWPArticulo = " & idArt
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        ReDim Preserve arrArtPla(UBound(arrArtPla) + 1)
        With arrArtPla(UBound(arrArtPla))
            .idArticulo = idArt
            .idPlaIntra = RsAux(1)
            .idPlaWeb = RsAux(2)
        End With
    End If
    RsAux.Close
End Sub

Private Sub AumentoProgressParcial()
On Error Resume Next
    If pbParcial.Value < pbParcial.Max Then pbParcial.Value = pbParcial.Value + 1
    pbParcial.Refresh
End Sub

Private Function acc_PasoPreciosACombo() As String
On Error GoTo errPPAC
Dim rsPC As rdoResultset, rsAP As rdoResultset
Dim lTCuota As Long
Dim iCont As Integer, iContArt As Integer

    acc_PasoPreciosACombo = ""
    
    frm_SetDetalleLogo "Pasando combos ..."
    
    Cons = "Select * From Articulo, Presupuesto Where ArtEsCombo = 1 And ArtEnUso = 1 " _
        & " And PreArtCombo = ArtID And PreEsPresupuesto = 0 And PreHabilitado = 1"
    
    Set rsPC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not rsPC.EOF
                
        iContArt = 0
        
        Cons = "Select count(*) From Presupuesto, PresupuestoArticulo " & _
            " Where PreArtCombo = " & rsPC!ArtID & " And PreID = PArPresupuesto "
        
        Set rsAP = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        iContArt = rsAP(0)
        rsAP.Close
               
        ReDim arrArtPre(0)
        
        If iContArt > 0 Then
            
            frm_SetDetalleLogo "Pasando ..."
            
            Cons = "Select PViTipoCuota, MAx(PlaNombre) as PlaNombre, TCuCantidad, TCuAbreviacion, TCuVencimientoC, sum((PViPrecio * PArCantidad)) as Precio, Count(*) as Cant " & _
                " From Presupuesto, PresupuestoArticulo, PrecioVigente, TipoCuota, TipoPlan " & _
                " Where PreArtCombo = " & rsPC!ArtID & " And PreID = PArPresupuesto And PViArticulo = PArArticulo " & _
                " And PViHabilitado <> 0  And PViMoneda = 1 And PViTipoCuota = TCuCodigo And PViPlan = PlaCodigo " & _
                " Group By PViTipoCuota, TCuCAntidad, TcuAbreviacion, TCuVencimientoC " & _
                " Order by PViTipoCuota"
            
            Set rsAP = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
            If Not rsAP.EOF Then
                'Borro los precios para este artículo
                Cons = "Delete * From PrecioVigente Where PViArticulo = " & rsPC!ArtID
                cAccess.Execute (Cons)
                
                'Válido que el artículo este en la tabla listaarticulo.
                Cons = "Select * From ListaArticulos Where LAArtID = " & rsPC!ArtID
                Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If rsAcc.EOF Then
                    rsAP.MoveLast: rsAP.MoveNext
                End If
                rsAcc.Close
            End If

            Do While Not rsAP.EOF
                
                If rsAP!Cant = iContArt Then
                    
                    Cons = "Select * From PrecioVigente Where PViArticulo = " & rsPC!ArtID
                    Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    rsAcc.AddNew
                    rsAcc("PViArticulo") = rsPC("ArtID")
                    rsAcc("PViTipoCuota") = rsAP("PViTipoCuota")
                    rsAcc("PViCantCuota") = rsAP("TCuCantidad")
                    rsAcc("PViDescripcion") = rsAP("TCuAbreviacion")
                    'El precio que guardo es el precio de la cuota.
                    rsAcc("PViPrecio") = Format(rsAP("Precio") / rsAP("TCuCantidad"), "###0")
                    rsAcc.Update
                    rsAcc.Close
                    
                    If rsAP("TCuVencimientoC") = 0 And rsAP("TCuCantidad") > 0 Then
                        InsertoEnArray rsPC("Artid"), rsAP("Precio") / rsAP("TCuCantidad"), rsAP("PlaNombre"), Trim(rsAP("TCuAbreviacion")), rsAP("PViTipoCuota")
                    End If
'                Else
'                    MsgBox "Combo " & Trim(rsPC!Artnombre) & " tipo de cuota " & rsAP("TCuAbreviacion") & " Retorno " & rsAP!Cant & " Filas"
                End If
                rsAP.MoveNext
            Loop
            rsAP.Close
            
            'Actualizo el precio en la tabla ListaArticulos
            For iCont = 1 To UBound(arrArtPre)
                Cons = "Select * From ListaArticulos Where LAArtID = " & arrArtPre(iCont).idArt
                Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not rsAcc.EOF Then
                    rsAcc.Edit
                    rsAcc("LAContado") = arrArtPre(iCont).Ctdo
                    rsAcc("LAPlan") = Trim(arrArtPre(iCont).Plan)
                    If arrArtPre(iCont).CuotaFinanciado = -1 Then
                        rsAcc("LAFinanciado") = Null
                        rsAcc("LATextoFinanciado") = Null
                    Else
                        rsAcc("LAFinanciado") = arrArtPre(iCont).CuotaFinanciado
                        rsAcc("LATextoFinanciado") = arrArtPre(iCont).TCuotaAbrev
                    End If
                    rsAcc.Update
                End If
                rsAcc.Close
            Next iCont
            
        End If
        rsPC.MoveNext
    Loop
    rsPC.Close
    Exit Function
    
errPPAC:
    acc_PasoPreciosACombo = f_SetError("Precios Combos")
End Function

Private Sub Label1_Click()
On Error Resume Next
    With tArticulo
        .SelStart = 0: .SelLength = Len(.Text)
        .SetFocus
    End With
End Sub

Private Sub opHacer_Click(Index As Integer)
        
    chEspecie.Enabled = (Index = 1)
    chTipo.Enabled = (Index = 1)
    chMarca.Enabled = (Index = 1)
    chGrupo.Enabled = (Index = 1)
    chPlanes.Enabled = (Index = 1)
    chGlosario.Enabled = (Index = 1)
    chStock.Enabled = (Index = 1)
    chPrecio.Enabled = (Index = 1)
    chLista.Enabled = (Index = 1)
    
    tArticulo.Enabled = (Index = 2)
    If Index = 2 Then
        tArticulo.BackColor = vbWindowBackground
    Else
        tArticulo.BackColor = vbButtonFace
        tArticulo.Text = ""
    End If
    
    If opHacer(0).Value Then
        frm_SetDetalleLogo "Sí o sí artículos, stock y precios, más otros cambios"
    ElseIf opHacer(1).Value Then
        frm_SetDetalleLogo "Artículos modificados y los pasos seleccionados"
    Else
        frm_SetDetalleLogo "Sólo el artículo, sus precios, stock y páginas"
    End If
    
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = ""
End Sub

Private Sub tArticulo_GotFocus()
    With tArticulo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tArticulo_KeyPress(KeyAscii As Integer)
On Error Resume Next
    
    If KeyAscii = vbKeyReturn Then
        If Val(tArticulo.Tag) > 0 Then
            bActualizar.SetFocus
        ElseIf Trim(tArticulo.Text) <> "" Then
            Cons = "Select Artid, ArtCodigo as Codigo,  ArtNombre as Nombre From Articulo"
            If IsNumeric(tArticulo.Text) Then
                Cons = Cons & " Where ArtCodigo = " & Val(tArticulo.Text) & _
                        " And Lower(ArtHabilitado) = 's' And ArtEnUso = 1"
            Else
                Cons = Cons & " Where ArtNombre like " & f_SetFilterFind(tArticulo.Text) & _
                    " And Lower(ArtHabilitado) = 's' And ArtEnUso = 1" & _
                    "Order by ArtNombre"
            End If
            db_FindCodigoNombre Cons, tArticulo, "Artículos"
            If Val(tArticulo.Tag) > 0 Then bActualizar.SetFocus
        End If
    End If

End Sub

Private Sub db_ShowHelpList(ByVal sQuery As String, ByVal textB As TextBox, ByVal sTitulo As String, Optional iColHide As Integer = 1)
On Error GoTo errSH
    
    Dim objHelp As New clsListadeAyuda
    With objHelp
        If .ActivarAyuda(cBase, sQuery, , iColHide, sTitulo) > 0 Then
            textB.Text = .RetornoDatoSeleccionado(2)
            textB.Tag = .RetornoDatoSeleccionado(0)
        End If
    End With
    Set objHelp = Nothing
    Exit Sub
errSH:
    clsGeneral.OcurrioError "Error al activar la lista de ayuda.", Err.Description
End Sub

Private Sub db_FindCodigoNombre(ByVal sQuery As String, ByVal textB As TextBox, Optional sTitLista As String = "Ayuda")
On Error GoTo errFCN

    Screen.MousePointer = 11
    Set RsAux = cBase.OpenResultset(sQuery, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        RsAux.MoveNext
        If Not RsAux.EOF Then
            RsAux.Close
            'Lista de ayuda
            db_ShowHelpList Cons, textB, sTitLista, 1
        Else
            RsAux.MoveFirst
            With textB
                .Text = Trim(RsAux(2))
                .Tag = RsAux(0)
            End With
        End If
    Else
        MsgBox "No existe un artículo con el dato ingresado.", vbInformation, "Buscar " & sTitLista
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errFCN:
    clsGeneral.OcurrioError "Error al buscar.", Err.Description, "Búsqueda general"
    Screen.MousePointer = 0
End Sub

Private Function GetRdoError() As String
On Error GoTo errGRE
Dim objRdoErr As rdoError
    GetRdoError = ""
    For Each objRdoErr In rdoErrors
        If GetRdoError <> "" Then GetRdoError = GetRdoError & vbCrLf
        GetRdoError = GetRdoError & objRdoErr.Description
    Next
    rdoErrors.Clear     'Clean collection
    Exit Function
errGRE:
    Resume Next
End Function

Private Sub frm_SetDetalleLogo(ByVal sMsg As String)
    lLogoDet.Caption = sMsg
    Me.Refresh
End Sub

Private Function f_SetError(ByVal sTitulo As String)
    f_SetError = String(30, ".") & sTitulo & vbCrLf & GetRdoError & vbCrLf & "Descripción: " & Err.Description & vbCrLf & String(30, ".") & sTitulo & vbCrLf
End Function

Private Function acc_InvocoArchivo() As String
On Error GoTo errIA
    EjecutarApp paFileInvoco
Exit Function
errIA:
    acc_InvocoArchivo = f_SetError("InvocoArchivo")
End Function

Private Function acc_ActualizarArticuloAccesorios() As String
On Error GoTo errPC
Dim sTexto As String
Dim lSum As Long
    
    acc_ActualizarArticuloAccesorios = ""
    If Not (chAccesorios.Value = 1 Or opHacer(0).Value) Then Exit Function
    
    lLogoPaso.Caption = "ArticuloAccesorio."
    lLogoDet.Caption = "Validando ..."
    
    Cons = "Select Count(*), IsNull(Sum(Len(rTrim(ArAWhereSql))), 0),  IsNull(Sum(ArAHabilitado), 0) From ArticuloAccesorios"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux(0) = 0 Then
        RsAux.Close
        Cons = "Delete * From ArticuloAccesorios"
        cAccess.Execute (Cons)
        Exit Function
    Else
    
        'Cuento los que hay en la bd access.
        Cons = "Select Count(*), Sum(Len(Trim(ArAWhereSql))),  Sum(ArAHabilitado) From ArticuloAccesorios"
        Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If rsAcc(0) = RsAux(0) And IIf(IsNull(rsAcc(1)), 0, rsAcc(1)) = RsAux(1) And IIf(IsNull(rsAcc(2)), 0, rsAcc(2)) = RsAux(2) Then
            rsAcc.Close
            RsAux.Close
            Exit Function
        End If
        rsAcc.Close
        pbParcial.Max = RsAux(0) + 2
        RsAux.Close
    End If
    
    pbParcial.Value = 1
    Cons = "Delete * From ArticuloAccesorios"
    cAccess.Execute (Cons)
    pbParcial.Value = 2
    
    lLogoDet.Caption = "Pasando datos ..."
    
    Cons = "Select * From ArticuloAccesorios"
    Set rsAcc = cAccess.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
    Cons = "Select * From ArticuloAccesorios"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        rsAcc.AddNew
        rsAcc("ArACodigo") = RsAux("ArACodigo")
        rsAcc("ArANombre") = RsAux("ArANombre")
        If Not IsNull(RsAux("ArAWhereSQL")) Then
            rsAcc("ArAWhereSQL") = RsAux("ArAWhereSQL")
        End If
        rsAcc("ArAHabilitado") = False
        If Not IsNull(RsAux("ArAHabilitado")) Then
            If RsAux("ArAHabilitado") > 0 Then
                rsAcc("ArAHabilitado") = True
            End If
        End If
        rsAcc.Update
        
        RsAux.MoveNext
        AumentoProgressParcial
    Loop
    RsAux.Close
    pbParcial.Value = 0
    Exit Function
    
errPC:
    acc_ActualizarArticuloAccesorios = f_SetError("ArticuloAccesorios")
End Function

