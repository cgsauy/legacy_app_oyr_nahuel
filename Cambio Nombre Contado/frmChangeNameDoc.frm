VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5EA2D00A-68AC-4888-98E6-53F6035BBEE3}#1.3#0"; "CGSABuscarCliente.ocx"
Begin VB.Form frmChangeNameDoc 
   BackColor       =   &H00B48246&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar Cliente a Contados"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChangeNameDoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin prjBuscarCliente.ucBuscarCliente txtCliente 
      Height          =   285
      Left            =   960
      TabIndex        =   29
      Top             =   3840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      Text            =   "_.___.___-_"
      DocumentoCliente=   1
      QueryFind       =   "EXEC [dbo].[prg_BuscarCliente] 0, '', '', '', '', '', '[KeyQuery]', 0, 0, '', '', 7"
      KeyQuery        =   "[KeyQuery]"
      NeedCheckDigit  =   0   'False
   End
   Begin VB.CommandButton bPrint 
      BackColor       =   &H00DEC4B0&
      Caption         =   "&Impresora"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox tDirNew 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   10
      Top             =   4200
      Width           =   5235
   End
   Begin VB.TextBox tCliNew 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   7
      Top             =   3840
      Width           =   5235
   End
   Begin VB.ComboBox cDireccion 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4200
      Width           =   1635
   End
   Begin VB.CommandButton bClose 
      BackColor       =   &H00DEC4B0&
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6780
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton bEmitir 
      BackColor       =   &H00DEC4B0&
      Caption         =   "&Emitir"
      Height          =   375
      Left            =   5580
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4680
      Width           =   1095
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsArticulo 
      Height          =   1395
      Left            =   60
      TabIndex        =   5
      Top             =   1380
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   2461
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14599344
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   12582912
      GridColorFixed  =   12582912
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0   'False
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
   End
   Begin MSComctlLib.StatusBar sbLineHelp 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   5145
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox tSerie 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   317
      Left            =   3780
      MaxLength       =   1
      TabIndex        =   3
      Top             =   360
      Width           =   242
   End
   Begin VB.TextBox tNumero 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   317
      Left            =   4020
      MaxLength       =   7
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin AACombo99.AACombo cLocal 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      BackColor       =   16777215
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
   Begin VB.Label lPC 
      BackStyle       =   0  'Transparent
      Caption         =   "Contado:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   28
      Top             =   4800
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label lRucNew 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "21.025996.0012"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   4320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lResult 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   3120
      Width           =   7815
   End
   Begin VB.Label lTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00DEC4B0&
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   6200
      TabIndex        =   24
      Top             =   2760
      Width           =   1755
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Dirección:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label lRucCI 
      BackStyle       =   0  'Transparent
      Caption         =   "R.U.C.:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   675
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF5F0&
      Caption         =   "  Nuevo Cliente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   3480
      Width           =   8055
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Emisión:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   22
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lFDoc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "10-Dic-1998"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6660
      TabIndex        =   21
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label labDato1 
      BackStyle       =   0  'Transparent
      Caption         =   "R.U.C.:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   20
      Top             =   780
      Width           =   675
   End
   Begin VB.Label lRuc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "21.025996.0012"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   19
      Top             =   780
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   780
      Width           =   735
   End
   Begin VB.Label lCliOrig 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Walter Adrián Occhiuzzi"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1020
      TabIndex        =   17
      Top             =   780
      Width           =   4575
   End
   Begin VB.Label lDirOrig 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Avda. Italia 2555"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1020
      TabIndex        =   16
      Top             =   1080
      Width           =   6915
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF5F0&
      Caption         =   "  Documento Contado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   8055
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "S&ucursal:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   795
   End
   Begin VB.Label lblNroDoc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Número:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3060
      TabIndex        =   2
      Top             =   360
      Width           =   675
   End
   Begin VB.Menu MnuCliente 
      Caption         =   "Cliente"
      Visible         =   0   'False
      Begin VB.Menu MnuCliNew 
         Caption         =   "Nuevo"
      End
      Begin VB.Menu MnuCliLine 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCliUpdate 
         Caption         =   "&Modificar"
      End
      Begin VB.Menu MnuCliFind 
         Caption         =   "Buscar"
      End
   End
End
Attribute VB_Name = "frmChangeNameDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Modificaciones
'   21-10-03        Veo si el nuevo ctdo lleva o no cofis según el cliente
'                        Puse seteo de impresora.

'   02-02-04        Corregí grabar cofis en los renglones, updateaba nueavemente a cero. Copio la zona del documento anterior.
'   08-07-04        Si el envío no se recepcionó y es de vta. telefónica le cambio el id del documento a la misma.
'   03-08-07        Elimine el cofis. Cuando cargo los artículos calculo nuevamente el iva y cuando copio el documento tomo el valor total del iva de la suma de la grilla.
'   15-09-07        Multiplico por la Q de arts el IVA que paso al documento
'......................................................................................................................
'Controles especiales.
'    1) Nota: le pongo cantidad a retirar cero a los renglones de la misma xq si la anulan pueden nuevamente retirar los art.

Private Type tDocumento
    ID As Long
    Serie As String
    Numero As Long
End Type

Dim oCnfgPrint As New clsImpresoraTicketsCnfg
Private Departamento As String, Localidad As String
Private DepartamentoNota As String, LocalidadNota As String

Private prmEFacturaProductivo As String

Private lDirFact As Long
Private sRemito As String
Private sEnvio As String, sVtaTelef As String
Private sTextRetira As String
Private sInst As String

Private Sub bClose_Click()
    Unload Me
End Sub

Private Sub bEmitir_Click()

    If txtCliente.Cliente.Codigo = 0 Then
        MsgBox "No hay un cliente seleccionado.", vbExclamation, "ATENCIÓN"
        txtCliente.SetFocus
        Exit Sub
    End If
    If Val(tNumero.Tag) = 0 Then
        MsgBox "No hay un contado seleccionado.", vbExclamation, "ATENCIÓN"
        cLocal.SetFocus
        Exit Sub
    End If
    
    If sVtaTelef <> "" Then MsgBox "Si el envío no fue recepcionado el pendiente se asociará a este nuevo documento.", vbInformation, "Atención"
    
    If MsgBox("¿Confirma emitir el nuevo documento?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
        act_Save
    End If

End Sub

Private Sub bPrint_Click()
On Error Resume Next
    
    prj_LoadConfigPrint True
    
    lPC.Caption = "Imp. Ctdo y Nota: " & paINContadoN
    If Not paPrintEsXDefNC Then lPC.ForeColor = &HC0& Else lPC.ForeColor = vbWhite

End Sub

Private Sub CargoDatosDireccion()
On Error GoTo errCargar
    Screen.MousePointer = 11
    tDirNew.Text = ""
    Departamento = ""
    Localidad = ""
    If cDireccion.ListIndex > -1 Then
        tDirNew.Text = " " & clsGeneral.ArmoDireccionEnTexto(cBase, cDireccion.ItemData(cDireccion.ListIndex))
        CargoDepartamentoLocalidad cDireccion.ItemData(cDireccion.ListIndex), Departamento, Localidad
    End If
    
errCargar:
    Screen.MousePointer = 0
End Sub

Private Sub cDireccion_Click()
On Error GoTo errCargar
    tDirNew.Text = ""
    If cDireccion.ListIndex <> -1 Then
        CargoDatosDireccion
    End If
errCargar:
    Screen.MousePointer = 0
End Sub

Private Sub cDireccion_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then tDirNew.SetFocus
End Sub

Private Sub cLocal_Change()
    LimpioDocumento
End Sub

Private Sub cLocal_Click()
    LimpioDocumento
End Sub

Private Sub cLocal_GotFocus()
    With cLocal
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    LineHelp "Seleccione la sucursal de emisión del documento."
End Sub

Private Sub cLocal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cLocal.ListIndex > -1 Then
            tSerie.SetFocus
        Else
            MsgBox "El ingreso del local es obligatorio.", vbExclamation, "ATENCIÓN"
        End If
    End If
End Sub
Private Sub cLocal_LostFocus()
    LineHelp ""
End Sub

Private Sub Form_Load()
    
    tDirNew.Tag = 2
    lPC.Caption = "Imp. Ctdo y Nota: " & paINContadoN
    
    Set txtCliente.Connect = cBase
    txtCliente.NeedCheckDigit = True
    
    InitGrid
    LimpioDocumento
    LimpioCliente
    
    oCnfgPrint.CargarConfiguracion "FacturaContado", "TicketContado"
    
    Cons = "Select SucCodigo, SucAbreviacion From Sucursal Order by SucAbreviacion"
    CargoCombo Cons, cLocal, ""
    
    If Not ValidarVersionEFactura Then
        MsgBox "La versión del componente CGSAEFactura está desactualizado, debe distribuir software." _
                    & vbCrLf & vbCrLf & "Se cancelará la ejecución.", vbCritical, "EFactura"
        End
    End If
    
    crAbroEngine
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    crCierroEngine
End Sub

Private Function ValidarVersionEFactura() As Boolean
On Error GoTo errEC
    With New clsCGSAEFactura
        ValidarVersionEFactura = .ValidarVersion()
    End With
    Exit Function
errEC:
End Function

Private Sub Label12_Click()
On Error Resume Next
    With cLocal
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub lRucCI_Click()
On Error Resume Next
    With txtCliente
        .SetFocus
    End With
End Sub

Private Sub tCliNew_GotFocus()
    With tCliNew
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tCliNew_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cDireccion.SetFocus
End Sub

Private Sub tDirNew_GotFocus()
    With tDirNew
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tDirNew_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then txtCliente.SetFocus
End Sub

Private Sub tNumero_Change()
    LimpioDocumento
End Sub

Private Sub tNumero_GotFocus()
   With tNumero
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    LineHelp " Ingrese el número del documento. ([Enter] buscar)"
End Sub

Private Sub tNumero_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Val(tNumero.Tag) > 0 Then
            txtCliente.SetFocus
        Else
            If Not IsNumeric(tNumero.Text) Then
                txtCliente.SetFocus
                Exit Sub
            End If
            If cLocal.ListIndex = -1 Then
                MsgBox "Seleccione un local.", vbExclamation, "ATENCIÓN"
                cLocal.SetFocus
                Exit Sub
            End If
            If tSerie.Text = "" Then
                MsgBox "Ingrese la serie del documento.", vbExclamation, "ATENCIÓN"
                tSerie.SetFocus
                Exit Sub
            End If
            FindDocumento
            If Val(tNumero.Tag) > 0 Then txtCliente.SetFocus
            If txtCliente.Cliente.Codigo > 0 Then bEmitir.Enabled = True
        End If
    End If
    
End Sub

Private Sub tNumero_LostFocus()
    LineHelp ""
End Sub

Private Sub tSerie_Change()
    LimpioDocumento
End Sub

Private Sub tSerie_GotFocus()
    With tSerie
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    LineHelp " Ingrese la serie del documento."
End Sub

Private Sub tSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        tSerie.Text = UCase(tSerie.Text)
        tNumero.SetFocus
    End If
End Sub

Private Sub tSerie_LostFocus()
    LineHelp vbNullString
End Sub

Private Sub LineHelp(ByVal sText As String)
    sbLineHelp.SimpleText = sText
End Sub

Private Sub InitGrid()
    With vsArticulo
        .Cols = 1
        .FormatString = ">Q|<Artículo|>A Retirar|>Unitario|>Importe|codigo|especifico"
        .ColWidth(0) = 450: .ColWidth(1) = 4000: .ColWidth(3) = 1200
        .ColHidden(5) = True
        .ColHidden(6) = True
    End With
End Sub

Private Sub LimpioDocumento()
    bEmitir.Enabled = False
    tNumero.Tag = "": lFDoc.Caption = "": lRuc.Caption = "": lFDoc.Tag = ""
   
    lCliOrig.Caption = "": lDirOrig.Caption = "": lCliOrig.Tag = "": lblNroDoc.Tag = "": labDato1.Tag = ""
    vsArticulo.Rows = 1: lTotal.Caption = "": lTotal.Tag = "": lRuc.Tag = ""
    lResult.Caption = ""
    sEnvio = ""
    sRemito = ""        'Guardo los remitos asociados al documento.
    sVtaTelef = ""
    sInst = ""
End Sub

Private Sub LimpioCliente()
    lDirFact = 0
    tCliNew.Text = "": tDirNew.Text = ""
    lRucNew.Caption = ""
    lRucNew.Tag = ""
    cDireccion.ListIndex = -1
    Departamento = ""
    Localidad = ""
    bEmitir.Enabled = False
    txtCliente_CambioTipoDocumento
End Sub

Private Function f_FindDocumento(ByVal sSerie As String, ByVal sNro As String, ByVal sucursal As Integer)
On Error GoTo errFD

    f_FindDocumento = 0
    Cons = "Select DocCodigo, DocFecha as Fecha, DocSerie as Serie, Convert(char(7),DocNumero) as Numero " & _
                " From Documento " & _
                " Where DocTipo IN (" & TipoDocumento.Contado & ")" & _
                " And DocSerie = '" & sSerie & "' And DocNumero = " & sNro & " And DocAnulado = 0 AND DocSucursal = " & sucursal
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        RsAux.MoveNext
        If Not RsAux.EOF Then
            Dim objHelp As New clsListadeAyuda
            With objHelp
                If .ActivarAyuda(cBase, Cons, 5000, 1, "Documentos") > 0 Then
                    f_FindDocumento = .RetornoDatoSeleccionado(0)
                End If
            End With
            Set objHelp = Nothing
        Else
            RsAux.MoveFirst
            f_FindDocumento = RsAux(0)
        End If
    End If
    RsAux.Close
    Exit Function

errFD:
    clsGeneral.OcurrioError "Error al buscar el documento.", Err.Description
End Function


Private Sub FindDocumento()
On Error GoTo errBD
Dim cAux As Currency, lCli As Long
    
    Screen.MousePointer = 11
    
    Dim cod As Long
    cod = f_FindDocumento(tSerie.Text, tNumero.Text, cLocal.ItemData(cLocal.ListIndex))
    
    Dim sCI As String

    Cons = "SELECT * " _
        & " FROM Documento INNER JOIN Renglon ON DocCodigo = RenDocumento " _
        & " INNER JOIN Articulo ON ArtID = RenArticulo " _
        & " INNER JOIN Moneda ON DocMoneda = MonCodigo " _
        & " INNER JOIN ArticuloFacturacion ON AFaArticulo = ArtID " _
        & " INNER JOIN TipoIva ON AFaIVA = IVACodigo" _
        & " INNER JOIN Cliente ON CliCodigo = DocCliente " _
        & " LEFT OUTER JOIN ArticuloEspecifico ON AEsDocumento = DocCodigo AND AEsArticulo = RenArticulo AND AEsTipoDocumento = 1" _
        & " WHERE DocCodigo = " & cod
        
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        RsAux.Close
        Screen.MousePointer = 0
        MsgBox "No se encontró un documento con esas características.", vbInformation, "ATENCIÓN"
    Else
        If RsAux!DocAnulado Then
            Screen.MousePointer = vbDefault
            RsAux.Close
            MsgBox "El documento seleccionado fue anulado, verifique.", vbExclamation, "ATENCIÓN"
        Else
            'Veo si el documento tiene alguna nota asociada.
            If Not DocumentInNota(RsAux!DocCodigo) Then
                
                If DocumentInFichaDev(RsAux!DocCodigo) Then
                    RsAux.Close
                    MsgBox "Existen fichas de devolución asociadas al contado, no podrá hacer cambios.", vbExclamation, "ATENCIÓN"
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                
                tNumero.Tag = RsAux!DocCodigo
                lFDoc.Caption = Format(RsAux!DocFecha, "d-Mmm-yyyy")
                lResult.Caption = " Se hará nota de devolución"
                lFDoc.Tag = RsAux!DocFModificacion
                
                lCli = RsAux!DocCliente
                If RsAux("CliTipo") = TipoCliente.Cliente And Not IsNull(RsAux("CliCiRUC")) Then sCI = RsAux("CliCIRUC")
                
                lTotal.Caption = Trim(RsAux!MonSigno) & " " & Format(RsAux!DocTotal, "#,##0.00") & " "
                lTotal.Tag = RsAux!DocMoneda & "|" & RsAux!DocTotal & "|" & RsAux!DocIVA
                If Not IsNull(RsAux!DocCofis) Then
                    lTotal.Tag = lTotal.Tag & "|" & RsAux!DocCofis
                Else
                    lTotal.Tag = lTotal.Tag & "|0"
                End If
                If Not IsNull(RsAux!DocVendedor) Then
                    lTotal.Tag = lTotal.Tag & "|" & RsAux!DocVendedor
                Else
                    lTotal.Tag = lTotal.Tag & "|0"
                End If
                If Not IsNull(RsAux!DocZona) Then
                    lTotal.Tag = lTotal.Tag & "|" & RsAux!DocZona
                Else
                    lTotal.Tag = lTotal.Tag & "|0"
                End If
                
                sTextRetira = "YA RETIRADO"
                Do While Not RsAux.EOF
                    With vsArticulo
                        .AddItem RsAux!RenCantidad
                        cAux = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = cAux
                        If IsNull(RsAux("AEsID")) Then
                            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!ArtCodigo, "(000,000)") & " " & Trim(RsAux!ArtNombre)
                            .Cell(flexcpText, .Rows - 1, 6) = 0
                        Else
                            .Cell(flexcpText, .Rows - 1, 1) = "Esp: " & RsAux("AEsID") & " " & Format(RsAux!ArtCodigo, "(000,000)") & " " & Trim(RsAux!ArtNombre)
                            .Cell(flexcpText, .Rows - 1, 6) = RsAux("AEsID")
                        End If
                        
                        .Cell(flexcpText, .Rows - 1, 5) = RsAux("ArtCodigo")
                        
                        cAux = RsAux("ArtTipo"): .Cell(flexcpData, .Rows - 1, 1) = cAux
                        
                        .Cell(flexcpData, .Rows - 1, 2) = Trim(RsAux!ArtNombre)
                        .Cell(flexcpText, .Rows - 1, 2) = RsAux!RenARetirar
                        If RsAux!RenARetirar > 0 Then
                            'Veo si no es del tipo servicio.
                            If Not EsTipoDeServicio(RsAux("ArtTipo")) Then sTextRetira = "POR RETIRAR"
                            'If paTipoArticuloServicio <> RsAux!ArtTipo Then sTextRetira = "POR RETIRAR"
                        End If
                        .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!RenPrecio, "#,##0.00")
                                    
                        cAux = RsAux!RenIVA: .Cell(flexcpData, .Rows - 1, 3) = cAux
                        
                        .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!RenPrecio * RsAux!RenCantidad, "#,##0.00")
                        
                        'If Not IsNull(RsAux!RenCofis) Then cAux = RsAux!RenCofis Else cAux = 0
                        cAux = RsAux("IVAPorcentaje"): .Cell(flexcpData, .Rows - 1, 4) = cAux
                    End With
                    RsAux.MoveNext
                Loop
                
                'Busco los envíos o remitos del contado.
                loc_DocumentRemito Val(tNumero.Tag)
                loc_DocumentEnvio Val(tNumero.Tag)
                loc_DocumentVtaTelefonica Val(tNumero.Tag)
                loc_DocumentoInstalacion Val(tNumero.Tag)
                
                If sEnvio <> "" Then lResult.Caption = lResult.Caption & ", hay envíos asociados "
                If sVtaTelef <> "" Then lResult.Caption = lResult.Caption & ", hay Ventas Telef. asociadas "
                If sRemito <> "" Then lResult.Caption = lResult.Caption & ", hay remitos asociados."
                If sInst <> "" Then lResult.Caption = lResult.Caption & ", hay instalaciones asociadas."
                
                If sTextRetira = "YA RETIRADO" And (sEnvio <> "" Or sRemito <> "") Then
                    loc_SetTextRetiraEnvioRetiro
                End If
                '........................................................
                
            Else
                RsAux.Close
                MsgBox "El contado esta asociado a notas de devolución, no puede efectuarle cambios.", vbExclamation, "ATENCIÓN"
                Screen.MousePointer = 0
                Exit Sub
            End If
            RsAux.Close
            
            lCliOrig.Tag = lCli
            'SetCustomerDoc lCli
            
            
            
            Cons = "SET QUOTED_IDENTIFIER ON SET CONCAT_NULL_YIELDS_NULL ON SET ANSI_PADDING ON SET ANSI_WARNINGS ON SET ANSI_NULLS ON SET ARITHABORT ON SELECT EComTipo, EcomXml.value('(/*:CFE//*:Encabezado/*:Receptor/*:TipoDocRecep)[1]', 'tinyint') TipoDoc, " & _
                    "EcomXml.value('(/*:CFE//*:Encabezado/*:Receptor/*:DocRecep)[1]', 'nvarchar(20)') Documento, " & _
                    "EcomXml.value('(/*:CFE//*:Encabezado/*:Receptor/*:RznSocRecep)[1]', 'nvarchar(100)') Nombre, " & _
                    "EcomXml.value('(/*:CFE//*:Encabezado/*:Receptor/*:DirRecep)[1]', 'nvarchar(100)') Direccion, " & _
                    "EcomXml.value('(/*:CFE//*:Encabezado/*:Receptor/*:CiudadRecep)[1]', 'nvarchar(20)') Localidad, " & _
                    "EcomXml.value('(/*:CFE//*:Encabezado/*:Receptor/*:DeptoRecep)[1]', 'nvarchar(20)') Departamento " & _
                    "FROM eComprobantes WHERE EComID = " & Val(tNumero.Tag)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                'Cargo los valores a partir del eComprobantes.
                If Not IsNull(RsAux("Nombre")) Then
                    lCliOrig.Caption = " " & Trim(RsAux!Nombre)
                End If
                lblNroDoc.Tag = RsAux("EComTipo")
                If Not IsNull(RsAux("Direccion")) Then lDirOrig.Caption = Trim(RsAux("Direccion"))

                If Not IsNull(RsAux("Departamento")) Then DepartamentoNota = Trim(RsAux("Departamento"))
                If Not IsNull(RsAux("Localidad")) Then LocalidadNota = Trim(RsAux("Localidad"))

                If Not IsNull(RsAux("TipoDoc")) Then lRuc.Tag = RsAux("TipoDoc")
                If Val(lRuc.Tag) = 2 And Not IsNull(RsAux("Documento")) Then
                    lRuc.Caption = RsAux("Documento")
                ElseIf Not IsNull(RsAux("Documento")) Then
                    labDato1.Tag = RsAux("Documento")
                Else
                    labDato1.Tag = sCI
                End If
                RsAux.Close
            Else
                RsAux.Close
                SetCustomerDoc lCli
            End If
            
            
            
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
errBD:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al buscar el documento.", Err.Description
End Sub

Private Sub SetCustomerDoc(ByVal lCodigo As Long)
Dim rs As rdoResultset
Dim lDirF As Long
    
    Cons = "SELECT CliCiRuc, CliTipo, PDDTipoDocIdentidad, CliDireccion, Nombre = CASE WHEN CliTipo = 1 THEN (RTrim(CPeApellido1) + RTrim(' ' + ISNULL(CPeApellido2, ''))+', ' + RTrim(CPeNombre1)) + RTrim(' ' + ISNULL(CPeNombre2,'')) ELSE RTRIM(CEmNombre) END, Ruc = CPeRuc " _
       & " FROM Cliente INNER JOIN PaisDelDocumento ON CliPaisDelDocumento = PDDId " _
       & " LEFT OUTER JOIN CPersona ON CliCodigo = CPeCliente " _
       & " LEFT OUTER JOIN CEmpresa ON CliCodigo = CEmCliente " _
       & " WHERE CliCodigo = " & lCodigo
    'Cons = GetQueryCustomer(lCodigo)
    Set rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not rs.EOF Then
        lCliOrig.Tag = lCodigo
        lCliOrig.Caption = " " & Trim(rs!Nombre)
        Dim idDirF As Long
        If Not IsNull(rs!CliDireccion) Then
            lDirF = loc_LoadDirFactura(lCodigo)
            If lDirF > 0 Then
                lDirOrig.Caption = " " & clsGeneral.ArmoDireccionEnTexto(cBase, lDirF)
                idDirF = lDirF
            Else
                lDirOrig.Caption = " " & clsGeneral.ArmoDireccionEnTexto(cBase, rs!CliDireccion)
                idDirF = rs("CliDireccion")
            End If
        Else
            lDirF = loc_LoadDirFactura(lCodigo)
            If lDirF > 0 Then lDirOrig.Caption = " " & clsGeneral.ArmoDireccionEnTexto(cBase, lDirF): idDirF = lDirF
        End If
        CargoDepartamentoLocalidad idDirF, DepartamentoNota, LocalidadNota
        
        If rs!CliTipo = TipoCliente.Empresa Then
            If Not IsNull(rs!CliCiRuc) Then lRuc.Caption = clsGeneral.RetornoFormatoRuc(Trim(rs!CliCiRuc))
        Else
            'labDato1.Caption = "C.I.:"
            If Not IsNull(rs!CliCiRuc) Then
                lCliOrig.Caption = Trim(lCliOrig.Caption) '& " (" & clsGeneral.RetornoFormatoCedula(rs!CliCiRuc) & ")"
                If Not IsNull(rs("CliCiRUC")) Then labDato1.Tag = rs("CliCiRUC")
            End If
            If Not IsNull(rs!Ruc) Then lRuc.Caption = clsGeneral.RetornoFormatoRuc(Trim(rs!Ruc))
        End If
        
        If lRuc.Caption <> "" Then
            lRuc.Tag = CGSA_TipoDocumentoDGI.TD_RUT
        Else
            lRuc.Tag = rs("PDDTipoDocIdentidad")
        End If
        
    End If
    rs.Close
    
End Sub

Private Sub SetNewCustomer(ByVal lCodigo As Long)
Dim rs As rdoResultset

    tCliNew.Text = " " & txtCliente.Cliente.Nombre
    
    Cons = "SELECT CliDireccion FROM Cliente WHERE CliCodigo = " & lCodigo
    Set rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not rs.EOF Then
        If Not IsNull(rs!CliDireccion) Then
            cDireccion.AddItem "Dirección Principal": cDireccion.ItemData(cDireccion.NewIndex) = rs!CliDireccion
            cDireccion.Tag = rs!CliDireccion
            cDireccion.ListIndex = 0
            lDirFact = rs!CliDireccion
        End If
        CargoDatosDireccion
        loc_LoadDirAuxiliares lCodigo
        If txtCliente.Cliente.Tipo = TipoCliente.Cliente Then
            If txtCliente.Cliente.RutPersona <> "" Then lRucNew.Caption = clsGeneral.RetornoFormatoRuc(Trim(txtCliente.Cliente.RutPersona))
        End If
        
    End If
    rs.Close
    If Val(tNumero.Tag) > 0 And txtCliente.Cliente.Codigo > 0 Then bEmitir.Enabled = True Else bEmitir = False

End Sub

Private Sub loc_DocumentRemito(ByVal lDoc As Long)
Dim rsLocal As rdoResultset
    
    sRemito = ""
    
    'Puedo tener un remito con mercadería pendiente de entrega (parciales).
    Cons = "Select * From Remito Where RemDocumento = " & lDoc
    Set rsLocal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    Do While Not rsLocal.EOF
        'Recorro la lista y asigno el remito.
        If sRemito <> "" Then sRemito = sRemito & ", "
        sRemito = sRemito & rsLocal("RemCodigo")
        rsLocal.MoveNext
    Loop
    rsLocal.Close
    
End Sub

Private Sub loc_DocumentEnvio(ByVal lDoc As Long)
Dim rsLocal As rdoResultset
    
    sEnvio = ""
    Cons = "SELECT EnvCodigo FROM Envio " & _
        " WHERE (EnvDocumento = " & lDoc & " OR EnvDocumento IN(SELECT RDoRemito FROM RemitoDocumento WHERE RDoDocumento = " & lDoc & "))" & _
        " And EnvTipo = 1"
    'Cons = "Select EnvCodigo From Envio Where EnvDocumento = " & lDoc & " And EnvTipo = 1"
    Set rsLocal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not rsLocal.EOF
        If sEnvio <> "" Then sEnvio = sEnvio & ", "
        sEnvio = sEnvio & rsLocal("EnvCodigo")
        rsLocal.MoveNext
    Loop
    rsLocal.Close
    
End Sub

Private Sub loc_DocumentVtaTelefonica(ByVal lDoc As Long)
Dim rsLocal As rdoResultset
    
    sVtaTelef = ""
    Cons = "Select VTeCodigo From VentaTelefonica Where VTeDocumento = " & lDoc
    Set rsLocal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not rsLocal.EOF
        If sVtaTelef <> "" Then sVtaTelef = sVtaTelef & ", "
        sVtaTelef = sVtaTelef & rsLocal("VTeCodigo")
        rsLocal.MoveNext
    Loop
    rsLocal.Close
    
End Sub

Private Sub loc_DocumentoInstalacion(ByVal lDoc As Long)
Dim rsR As rdoResultset
    Cons = "Select InsID From Instalacion Where InsTipoDocumento = 1 And InsDocumento = " & lDoc
    Set rsR = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsR.EOF
        If sInst <> "" Then sInst = sInst & ", "
        sInst = sInst & rsR("InsID")
        rsR.MoveNext
    Loop
    rsR.Close
End Sub

Private Sub loc_SetTextRetiraEnvioRetiro()
Dim rsLocal As rdoResultset
    If sEnvio <> "" Then
        Cons = "Select * From  RenglonEnvio, Articulo " & _
                    "Where RevEnvio IN (" & sEnvio & ") And RevAEntregar > 0 " & _
                    "And RevArticulo = ArtID And ArtTipo not in ( " & tTiposArtsServicio & ")"
        Set rsLocal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not rsLocal.EOF Then sTextRetira = "HAY ENVÍOS"
        rsLocal.Close
    End If
    If sRemito <> "" Then
        Cons = "Select * From  RenglonRemito, Articulo " & _
                    "Where RReRemito IN (" & sRemito & ") And RReAEntregar > 0 " & _
                    "And RReArticulo = ArtID And ArtTipo not in ( " & tTiposArtsServicio & ")"
        Set rsLocal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not rsLocal.EOF Then
            If sTextRetira = "HAY ENVÍOS" Then
                sTextRetira = sTextRetira & " y REMITOS"
            Else
                sTextRetira = "HAY REMITOS"
            End If
        End If
        rsLocal.Close
    End If
    
End Sub

Private Function DocumentInFichaDev(ByVal lDoc As Long) As Boolean
Dim rsFic As rdoResultset

    DocumentInFichaDev = False
    Cons = "Select * From Devolucion Where DevFactura = " & lDoc _
        & " And DevNota Is Null And DevLocal Is Not Null And DevFAltaLocal Is Not Null" _
        & " And DevAnulada Is Null"
    Set rsFic = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    If Not rsFic.EOF Then DocumentInFichaDev = True
    rsFic.Close
    
End Function

Private Function DocumentInNota(ByVal lDoc As Long) As Boolean
Dim rsNot As rdoResultset
    DocumentInNota = False
    Cons = "Select * From Nota, Documento Where NotFactura = " & lDoc _
        & " And DocAnulado = 0 And NotNota = DocCodigo"
    Set rsNot = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsNot.EOF Then DocumentInNota = True
    rsNot.Close
End Function

Private Sub act_Save()
Dim RsDoc As rdoResultset
Dim sSerie As String, lNroDoc As Long

    'Válido nuevamente que no me hayan insertado una nota.
    If DocumentInNota(Val(tNumero.Tag)) Then
        MsgBox "El documento posee una nota reciente, verifique.", vbCritical, "ATENCIÓN"
        LimpioDocumento
        txtCliente.Text = ""
        txtCliente.DocumentoCliente = DC_CI
        LimpioCliente
        Exit Sub
    End If
    
    If DocumentInFichaDev(Val(tNumero.Tag)) Then
        MsgBox "Al documento se le asignaron fichas de devolución, verifique.", vbCritical, "ATENCIÓN"
        LimpioDocumento
        txtCliente.Text = ""
        txtCliente.DocumentoCliente = DC_CI
        LimpioCliente
        Exit Sub
    End If
    
    
    'Veo dirección del nuevo cliente
    'Si paso la Validación controlo la direccion que factura-----------------------------------------------------------------
    If cDireccion.ListIndex <> -1 Then
        On Error Resume Next
        
        If lDirFact <> cDireccion.ItemData(cDireccion.ListIndex) Then        'Cambio Dir Facutua
            
            If MsgBox("Ud. cambió la dirección con la que el cliente factura habitualmente." & vbCrLf & "Quiere que esta dirección quede por defecto para facturar.", vbQuestion + vbYesNo, "Dirección por Defecto al Facturar") = vbYes Then
            
                If cDireccion.ItemData(cDireccion.ListIndex) <> Val(cDireccion.Tag) Then        'Dir. selecc. <> a la Ppal.
                    
                    Cons = "Select * from DireccionAuxiliar Where DAuCliente = " & txtCliente.Cliente.Codigo & " And DAuDireccion = " & cDireccion.ItemData(cDireccion.ListIndex)
                    Set RsDoc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If Not RsDoc.EOF Then
                        RsDoc.Edit: RsDoc!DAuFactura = True: RsDoc.Update
                    End If
                    RsDoc.Close
                    
                End If
                
                If lDirFact <> Val(cDireccion.Tag) Then      'La gDirFactura Anterior no era la ppal, la desmarco
                    Cons = "Select * from DireccionAuxiliar Where DAuCliente = " & txtCliente.Cliente.Codigo & " And DAuDireccion = " & lDirFact
                    Set RsDoc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If Not RsDoc.EOF Then
                        RsDoc.Edit: RsDoc!DAuFactura = False: RsDoc.Update
                    End If
                    RsDoc.Close
                End If
                lDirFact = cDireccion.ItemData(cDireccion.ListIndex)
                
            End If
        End If
    End If

    If prmEFacturaProductivo = "" Then CargareFacturaONOFF

    FechaDelServidor
    
    '.......................................................................................
    On Error GoTo errBT
    cBase.BeginTrans
    
    On Error GoTo errRB
    'Válido fecha de modificación del documento.
    'Si hay que anularlo ya lo hago.
    Cons = "Select * From Documento Where DocCodigo = " & Val(tNumero.Tag)
    Set RsDoc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If CDate(lFDoc.Tag) <> RsDoc!DocFModificacion Then
        cBase.RollbackTrans
        RsDoc.Close
        MsgBox "El documento fue modificado por otra terminal, verifique.", vbExclamation, "ATENCIÓN"
        LimpioDocumento
        txtCliente.Text = ""
        txtCliente.DocumentoCliente = DC_CI
        LimpioCliente
        Exit Sub
    End If
    
    'Inserto el Nuevo Contado y obtengo los prm que paso.
    Dim oCli As New clsClienteCFE
    With oCli
        .Codigo = txtCliente.Cliente.Codigo
        If txtCliente.Cliente.Tipo = TC_Empresa Then
            .RUT = txtCliente.Cliente.Documento
            .CodigoDGICI = TD_RUT
        Else
            .CI = clsGeneral.QuitoFormatoCedula(txtCliente.Cliente.Documento)
            If txtCliente.Cliente.RutPersona <> "" Then
                .CodigoDGICI = TD_RUT
                .RUT = txtCliente.Cliente.RutPersona
            Else
                .CodigoDGICI = txtCliente.Cliente.TipoDocumento.TipoDocIdDGI
            End If
        End If
        .NombreCliente = tCliNew.Text
        .Direccion.Departamento = Departamento
        .Direccion.Localidad = Localidad
        .Direccion.Domicilio = tDirNew.Text
    End With
    
    Dim caeNewDoc As clsCAEDocumento
    Dim oDocNew As New clsDocumentoCGSA
    Set caeNewDoc = save_CopyDocument(oDocNew)
    Set oDocNew.Cliente = oCli
    Set oDocNew.Conexion = cBase
    oDocNew.Codigo = oDocNew.InsertoDocumentoBD(0)
    '.......................................................................................
    
    RsDoc.Edit
'    'TODO://
'    RsDoc("DocAnulado") = 1
    RsDoc!DocFModificacion = Format(Now, "yyyy-mm-dd hh:mm:ss")
    RsDoc!DocComentario = "Cambio nombre doc (" & oDocNew.Codigo & ") " & caeNewDoc.Serie & " " & caeNewDoc.Numero
    RsDoc.Update
    RsDoc.Close
    
        'FIRMA NUEVO CONTADO.
    
    'le pongo la cantidad a retirar en cero.
    save_SetCantRenglonDocumento Val(tNumero.Tag)
    '.....................................................................................
    
    If sVtaTelef <> "" Then
        'Cambio la venta telefonica para el nuevo documento.
        Cons = "Select * From VentaTelefonica Where VTeDocumento = " & Val(tNumero.Tag)
        Set RsDoc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsDoc.EOF Then
            RsDoc.Edit
            RsDoc!VTeDocumento = oDocNew.Codigo
            RsDoc.Update
        End If
        RsDoc.Close
    End If
    
'GENERO LA NOTA
'TODO:
'If 1 = 2 Then
    Dim oCliNota As New clsClienteCFE
    With oCliNota
        If Val(lCliOrig.Tag) = 0 Then RsAux.Edit
        .Codigo = lCliOrig.Tag
        If lRuc.Caption <> "" Then
            .RUT = lRuc.Caption
            .CodigoDGICI = TD_RUT
        ElseIf labDato1.Tag <> "" And Val(lRuc.Tag) > 0 Then
            .CI = clsGeneral.QuitoFormatoCedula(labDato1.Tag)
            .CodigoDGICI = lRuc.Tag
        Else
            .CI = clsGeneral.QuitoFormatoCedula(labDato1.Tag)
        End If
        .NombreCliente = lCliOrig.Caption
        .Direccion.Departamento = DepartamentoNota
        .Direccion.Localidad = LocalidadNota
        .Direccion.Domicilio = lDirOrig.Caption
    End With
    
    Dim caeNota As clsCAEDocumento
    Dim oNota As New clsDocumentoCGSA
    Set oNota.Cliente = oCliNota
    Set caeNota = save_SetDocumentNota(oNota)
    Set oNota.Conexion = cBase
    oNota.Codigo = oNota.InsertoDocumentoBD(CLng(tNumero.Tag))
    
    Dim oDocRel As New clsDocumentoAsociado
    With oDocRel
        .Devuelve = oDocNew.Total
        .Fecha = gFechaServidor
        .Serie = tSerie.Text
        .Numero = tNumero.Text
        .Tipo = TD_Contado
        If Val(lblNroDoc.Tag) > 0 Then .TipoEFactura = Val(lblNroDoc.Tag)
    End With
    
    oNota.DocumentosAsociados.Add oDocRel
        
'End If

    If sEnvio <> "" Then
    
        'Tengo que seguir los siguientes pasos.
        '1) Cambiar el documento de los envíos que estén apuntando al documento.
        '2) Cambiar el documento en la tabla remitodocumento
        '3) Cambiar el documento de los específicos que apunten a este documento.
    
    
        'Modifico el documento en los envíos.
        Cons = "Select * from Envio Where EnvCodigo IN (" & sEnvio & ") AND EnvDocumento = " & Val(tNumero.Tag)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            RsAux.Edit
            RsAux("EnvDocumento") = oDocNew.Codigo
            If RsAux("EnvDocumentoFactura") = Val(tNumero.Tag) Then
                RsAux("EnvDocumentoFactura") = oDocNew.Codigo
            End If
            RsAux.Update
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        Cons = "UPDATE RemitoDocumento SET RDoDocumento = " & oDocNew.Codigo & " WHERE RDoDocumento = " & Val(tNumero.Tag)
        cBase.Execute Cons
                
    End If
    
    Cons = "UPDATE ArticuloEspecifico SET AEsDocumento = " & oDocNew.Codigo & " WHERE AEsDocumento = " & Val(tNumero.Tag) & " AND AEsTipoDocumento = 1"
    cBase.Execute Cons
    
    If sRemito <> "" Then
        Cons = "Select * From Remito Where RemCodigo IN (" & sRemito & ")"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            RsAux.Edit
            RsAux("RemDocumento") = oDocNew.Codigo
            RsAux.Update
            RsAux.MoveNext
        Loop
        RsAux.Close
    End If
    
    If sInst <> "" Then
    
        Cons = "Select * From Instalacion Where InsID IN (" & sInst & ")"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            RsAux.Edit
            RsAux("InsDocumento") = oDocNew.Codigo
            RsAux.Update
            RsAux.MoveNext
        Loop
        RsAux.Close
        
    End If
    
'    'Si hay documentos pendientes para el documento anterior le asigno el nuevo.
    Cons = "UPDATE DocumentoPendiente SET DPeDocumento = " & oDocNew.Codigo & _
        " WHERE DPeDocumento = " & Val(tNumero.Tag)
    cBase.Execute (Cons)
    
    'TODO: quitar anulación de texto.
    cBase.Execute "EXEC prg_PosInsertoDocumentosATickets '" & oDocNew.Codigo & "', " & oCnfgPrint.ImpresoraTickets
    cBase.Execute "EXEC prg_PosInsertoDocumentosATickets '" & oNota.Codigo & "', " & oCnfgPrint.ImpresoraTickets

    
    cBase.CommitTrans
    '.......................................................................................
    On Error GoTo errPrint
    Dim sCFE As String
    sCFE = EmitirCFE(oDocNew, caeNewDoc)
    If sCFE <> "" Then
        MsgBox "ATENCIÓN no se firmó el contado: " & sCFE, vbCritical, "ATENCIÓN"
    End If
    
    sCFE = ""
    sCFE = EmitirCFE(oNota, caeNota)
    If sCFE <> "" Then
        MsgBox "ATENCIÓN no se firmó la nota: " & sCFE, vbCritical, "ATENCIÓN"
    End If
    
    'Imprimo el nuevo contado.
'    If oDocNew.Codigo > 0 Then loc_PrintContado oDocNew.Codigo
'    If oNota.Codigo > 0 Then loc_PrintNota oNota.Codigo
    
    Set oNota = Nothing
    Set oDocNew = Nothing
    
    'Si hay nota la imprimo.
    On Error Resume Next
    
    LimpioDocumento
    LimpioCliente
    txtCliente.Text = ""
    cLocal.Text = "": tSerie.Text = "": tNumero.Text = ""
    Screen.MousePointer = 0
    Exit Sub
    
    
errPrint:
    Screen.MousePointer = 0
    MsgBox "Error al imprimir los documentos: " & Err.Description, vbCritical, "IMPRESIÓN"
    Exit Sub
errBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al iniciar la transacción.", Err.Description, "Grabar cambio de nombre a contado"
    Exit Sub
errRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al intentar almacenar la información.", Err.Description, "Grabar cambio de nombre a contado"
    Exit Sub
errRB:
    Resume errRoll
End Sub

Private Function save_CopyDocument(ByRef oDocNew As clsDocumentoCGSA) As clsCAEDocumento
Dim arrAux() As String
Dim iCount As Integer
Dim sNroDoc As String
Dim cIVA As Currency, cAux As Currency

    Dim tipoCAE As Byte
    tipoCAE = IIf((txtCliente.Cliente.Tipo = TC_Empresa And txtCliente.Cliente.Documento <> "") Or (txtCliente.Cliente.Tipo = TC_Persona And txtCliente.Cliente.RutPersona <> ""), CFE_eFactura, CFE_eTicket)
    
    Dim CAE As New clsCAEDocumento
    'Pido el Número de Documento-------------
    If prmEFacturaProductivo = "0" Then
        sNroDoc = NumeroDocumento(paDContado)
        With CAE
            .Desde = 1
            .Hasta = 9999999
            .Serie = Mid(sNroDoc, 1, 1)
            .Numero = CLng(Trim(Mid(sNroDoc, 2, Len(sNroDoc))))
            .IdDGI = "9014"
            .TipoCFE = tipoCAE
            .Vencimiento = "31/12/" & CStr(Year(Date))
        End With
    Else
        Dim caeG As New clsCAEGenerador
        Set CAE = caeG.ObtenerNumeroCAEDocumento(cBase, tipoCAE, paCodigoDGI)
        Set caeG = Nothing
    End If
    '----------------------------------------------------
    
    Set save_CopyDocument = CAE
    'Sumo el iva de los renglones.
    arrAux = Split(lTotal.Tag, "|")
    
    '3/8/2007       acá cambie y tomo los valores nuevos de la grilla
    With vsArticulo
        For iCount = 1 To .Rows - 1
            'tomo el iva del parámetro.
            cAux = .Cell(flexcpText, iCount, 3)
            cAux = Format((cAux - (cAux / (1 + (.Cell(flexcpData, iCount, 4) / 100)))) * CInt(.Cell(flexcpText, iCount, 0)), "#,##0.00")
            cIVA = cIVA + cAux
        Next
    End With
'    lNewDoc = save_InsertDoc(Contado, CAE.Serie, CAE.Numero, txtCliente.Cliente.Codigo, arrAux(0), arrAux(1), cIVA, 0, CLng(arrAux(4)), CLng(arrAux(5)))
    With oDocNew
        .Emision = gFechaServidor
        .Comentario = "Cambio nombre doc (" & tNumero.Tag & ") " & tSerie.Text & " " & tNumero.Text
        .Digitador = miConexion.UsuarioLogueado(True)
        .IVA = cIVA
        .Total = arrAux(1)
        .Moneda.Codigo = arrAux(0)
        .Numero = CAE.Numero
        .Serie = CAE.Serie
        .sucursal = paCodigoDeSucursal
        .Tipo = TD_Contado
        .Vendedor = CLng(arrAux(4))
    End With

    Dim oRenglon As clsDocumentoRenglon
    With vsArticulo
        For iCount = 1 To .Rows - 1
            cAux = .Cell(flexcpText, iCount, 3)
            cAux = Format(cAux - (cAux / (1 + (.Cell(flexcpData, iCount, 4) / 100))), "#,##0.00")
            
            Set oRenglon = New clsDocumentoRenglon
            oDocNew.Renglones.Add oRenglon
            
            oRenglon.Articulo.ID = .Cell(flexcpData, iCount, 0)
            oRenglon.Articulo.Nombre = .Cell(flexcpData, iCount, 2)
            oRenglon.Articulo.Codigo = .Cell(flexcpText, iCount, 5)
            oRenglon.Articulo.IDEspecifico = Val(.Cell(flexcpText, iCount, 6))
            oRenglon.Articulo.TipoIVA.Porcentaje = .Cell(flexcpData, iCount, 4)
            oRenglon.Articulo.TipoArticulo = .Cell(flexcpData, iCount, 1)
            oRenglon.Cantidad = .Cell(flexcpText, iCount, 0)
            oRenglon.CantidadARetirar = .Cell(flexcpText, iCount, 2)
            oRenglon.IVA = cAux
            oRenglon.Precio = CCur(.Cell(flexcpText, iCount, 3))
        
        Next
    End With
    
End Function

Private Function save_InsertDoc(ByVal eTipo As TipoDocumento, ByVal sSerie As String, ByVal lNro As Long, _
                                        ByVal lCliente As Long, ByVal lMoneda As Long, ByVal cTotal As Currency, ByVal cIVA As Currency, _
                                        ByVal cCofis As Currency, Optional lVendedor As Long = 0, Optional lZona As Long = 0) As Long

    Cons = "Select * From Documento Where DocCodigo = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.AddNew
    RsAux!DocFecha = Format(Now, "yyyy-mm-dd hh:nn:ss")
    RsAux!DocFModificacion = RsAux!DocFecha
    RsAux!DocTipo = eTipo
    RsAux!DocSerie = sSerie
    RsAux!DocNumero = lNro
    RsAux!DocCliente = lCliente
    RsAux!DocMoneda = lMoneda
    RsAux!DocTotal = cTotal
    RsAux!DocIVA = cIVA
'    RsAux!DocCofis = cCofis
    RsAux!DocAnulado = 0
    RsAux!DocSucursal = paCodigoDeSucursal
    RsAux!DocUsuario = miConexion.UsuarioLogueado(True)
    If lVendedor > 0 Then RsAux!DocVendedor = lVendedor
    If lZona > 0 Then RsAux!DocZona = lZona
    If TipoDocumento.Contado = eTipo Then RsAux!DocComentario = "Cambio nombre doc (" & tNumero.Tag & ") " & tSerie.Text & " " & tNumero.Text
    RsAux.Update
    RsAux.Close
    
    Cons = "Select Max(DocCodigo) From Documento Where DocTipo = " & eTipo _
        & " And DocSerie = '" & sSerie & "' And DocNumero = " & lNro & " And DocCliente  = " & lCliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    save_InsertDoc = RsAux(0)
    RsAux.Close
    
End Function

Private Function save_SetDocumentNota(ByRef oNota As clsDocumentoCGSA) As clsCAEDocumento
Dim arrAux() As String
Dim iCount As Integer
Dim sNroDoc As String
Dim sS As String, lN As Long

    Dim tipoCAE As Byte
    If Val(lblNroDoc.Tag) > 0 Then
        tipoCAE = IIf(Val(lblNroDoc.Tag) = CFE_eFactura, CGSA_TiposCFE.CFE_eFacturaNotaCredito, CGSA_TiposCFE.CFE_eTicketNotaCredito)
    Else
        tipoCAE = IIf(lRuc.Caption <> "", CGSA_TiposCFE.CFE_eFacturaNotaCredito, CGSA_TiposCFE.CFE_eTicketNotaCredito)
    End If
    
    Dim CAE As New clsCAEDocumento
    If prmEFacturaProductivo = "0" Then
        'Pido el Número de Documento-------------
        sNroDoc = NumeroDocumento(paDNDevolucion)
        sS = Mid(sNroDoc, 1, 1)
        lN = CLng(Trim(Mid(sNroDoc, 2, Len(sNroDoc))))
        '----------------------------------------------------
        With CAE
            .Desde = 1
            .Hasta = 9999999
            .Serie = sS
            .Numero = lN
            .IdDGI = "901412"
            .TipoCFE = tipoCAE
            .Vencimiento = "31/12/" & CStr(Year(Date))
        End With
    Else
        Dim caeG As New clsCAEGenerador
        Set CAE = caeG.ObtenerNumeroCAEDocumento(cBase, tipoCAE, paCodigoDGI)
        Set caeG = Nothing
    End If
    
    arrAux = Split(lTotal.Tag, "|")
'    lNota = save_InsertDoc(NotaDevolucion, CAE.Serie, CAE.Numero, lCliOrig.Tag, arrAux(0), arrAux(1), arrAux(2), arrAux(3), CLng(arrAux(4)))
    With oNota
        .Emision = gFechaServidor
        .Comentario = "Cambio nombre doc (" & tNumero.Tag & ") " & tSerie.Text & " " & tNumero.Text
        .Digitador = miConexion.UsuarioLogueado(True)
        .IVA = arrAux(2)
        .Total = arrAux(1)
        .Moneda.Codigo = arrAux(0)
        .Numero = CAE.Numero
        .Serie = CAE.Serie
        .sucursal = paCodigoDeSucursal
        .Tipo = TD_NotaDevolucion
        .Vendedor = CLng(arrAux(4))
        .NotaDevuelve = .Total
    End With
    Set save_SetDocumentNota = CAE
    
    Dim oRenglon As clsDocumentoRenglon
    With vsArticulo
        For iCount = 1 To .Rows - 1
            Set oRenglon = New clsDocumentoRenglon
            oNota.Renglones.Add oRenglon
            oRenglon.Articulo.ID = .Cell(flexcpData, iCount, 0)
            oRenglon.Articulo.Nombre = .Cell(flexcpData, iCount, 2)
            oRenglon.Articulo.Codigo = .Cell(flexcpText, iCount, 5)
            oRenglon.Articulo.IDEspecifico = Val(.Cell(flexcpText, iCount, 6))
            oRenglon.Articulo.TipoIVA.Porcentaje = .Cell(flexcpData, iCount, 4)
            oRenglon.Articulo.TipoArticulo = .Cell(flexcpData, iCount, 1)
            oRenglon.Cantidad = .Cell(flexcpText, iCount, 0)
            oRenglon.CantidadARetirar = 0
            oRenglon.IVA = .Cell(flexcpData, iCount, 3)
            oRenglon.Precio = CCur(.Cell(flexcpText, iCount, 3))
        Next
    End With

End Function

Private Sub save_SetCantRenglonDocumento(ByVal lDoc As Long)
Dim rsRen As rdoResultset
Dim iCount As Integer

    With vsArticulo
        Cons = ""
        For iCount = 1 To .Rows - 1
            If Val(.Cell(flexcpText, iCount, 2)) > 0 Then
                If Cons <> "" Then Cons = Cons & ","
                Cons = Cons & .Cell(flexcpData, iCount, 0)
            End If
        Next
        If Cons <> "" Then
            Cons = "UPDATE Renglon SET RenARetirar = 0 WHERE RenDocumento = " & lDoc & " AND RenArticulo IN (" & Cons & ")"
            cBase.Execute Cons
        End If
        
    End With
    
End Sub

Private Sub loc_PrintNota(ByVal lCodDoc As Long)
On Error GoTo ErrCrystal
Dim Result As Integer, JobSRep1 As Integer, JobSRep2 As Integer, jobnum As Integer
Dim NombreFormula As String, CantForm As Integer, aTexto As String
Dim sNomMoneda As String, sSingMoneda As String, lCodMoneda As Long

    Screen.MousePointer = 11
    'Inicializo el Reporte y SubReportes
    jobnum = crAbroReporte(gPathListados & "NotaDevolucion.RPT")
    If jobnum = 0 Then GoTo ErrCrystal
    
    'Configuro la Impresora
    If Trim(Printer.DeviceName) <> Trim(paINContadoN) Then SeteoImpresoraPorDefecto paINContadoN
    If Not crSeteoImpresora(jobnum, Printer, paINContadoB) Then GoTo ErrCrystal
    
    'Obtengo la cantidad de formulas que tiene el reporte.
    CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
    If CantForm = -1 Then GoTo ErrCrystal
    
    lCodMoneda = Mid(lTotal.Tag, 1, InStr(1, lTotal.Tag, "|") - 1)
    BuscoDatosMoneda lCodMoneda, sSingMoneda, sNomMoneda
    
    'Cargo Propiedades para el reporte Contado --------------------------------------------------------------------------------
    For I = 0 To CantForm - 1
        NombreFormula = crObtengoNombreFormula(jobnum, I)
        
        Select Case LCase(NombreFormula)
            Case "": GoTo ErrCrystal
            Case "nombredocumento": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & paDNDevolucion & "'")
            Case "cliente"
                Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lCliOrig.Caption) & "'")
            Case "direccion": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lDirOrig.Caption) & "'")
            Case "ruc"
                If lRuc.Caption <> "" Then Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lRuc.Caption) & "'")
            Case "codigobarras": Result = crSeteoFormula(jobnum%, NombreFormula, "''")
            Case "signomoneda": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & sSingMoneda & "'")
            Case "nombremoneda": Result = crSeteoFormula(jobnum%, NombreFormula, "'(" & sNomMoneda & ")'")
            Case "textoretira"
                Result = crSeteoFormula(jobnum%, NombreFormula, "''")
            Case Else: Result = 1
        End Select
        If Result = 0 Then GoTo ErrCrystal
    Next
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    'Seteo la Query del reporte-----------------------------------------------------------------
    Cons = "SELECT Documento.DocCodigo , Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal, Documento.DocIVA, Documento.DocVendedor" _
            & " From " & paBD & ".dbo.Documento Documento " _
            & " Where DocCodigo = " & lCodDoc
    If crSeteoSqlQuery(jobnum%, Cons) = 0 Then GoTo ErrCrystal
        
    'Subreporte srContado.rpt  y srContado.rpt - 01-----------------------------------------------------------------------------
    JobSRep1 = crAbroSubreporte(jobnum, "srContado.rpt")
    If JobSRep1 = 0 Then GoTo ErrCrystal
    
    Cons = "SELECT Renglon.RenDocumento, Renglon.RenCantidad, Renglon.RenPrecio, Renglon.RenDescripcion," _
            & " From { oj " & paBD & ".dbo.Renglon Renglon INNER JOIN " _
                           & paBD & ".dbo.Articulo Articulo ON Renglon.RenArticulo = Articulo.ArtId}"
    If crSeteoSqlQuery(JobSRep1, Cons) = 0 Then GoTo ErrCrystal
    
    JobSRep2 = crAbroSubreporte(jobnum, "srContado.rpt - 01")
    If JobSRep2 = 0 Then GoTo ErrCrystal
    If crSeteoSqlQuery(JobSRep2, Cons) = 0 Then GoTo ErrCrystal
    '-------------------------------------------------------------------------------------------------------------------------------------
    
    'If crMandoAPantalla(JobNum, "Factura Contado") = 0 Then GoTo ErrCrystal
    If crMandoAImpresora(jobnum, 1) = 0 Then GoTo ErrCrystal
    If Not crInicioImpresion(jobnum, True, False) Then GoTo ErrCrystal
    
    If Not crCierroSubReporte(JobSRep1) Then GoTo ErrCrystal
    If Not crCierroSubReporte(JobSRep2) Then GoTo ErrCrystal
    
    'crEsperoCierreReportePantalla
    Screen.MousePointer = 0
    Exit Sub

ErrCrystal:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError crMsgErr, Err.Description, "Impresión Nota"
    On Error Resume Next
    Screen.MousePointer = 11
    crCierroSubReporte JobSRep1
    crCierroSubReporte JobSRep2
    Screen.MousePointer = 0
    Exit Sub
End Sub

Private Sub BuscoDatosMoneda(ByVal Codigo As Long, sSing As String, sNombre As String)

    On Error GoTo ErrBU
    Dim rs As rdoResultset

    Cons = "SELECT * FROM Moneda WHERE MonCodigo = " & Codigo
    Set rs = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rs.EOF Then
        sNombre = Trim(rs!MonNombre)
        sSing = Trim(rs!MonSigno)
    End If
    rs.Close
    Exit Sub
    
ErrBU:
End Sub

Private Sub loc_LoadDirAuxiliares(ByVal aIdCliente As Long)

    On Error GoTo errCDA
    Dim rsDA As rdoResultset
    
    Cons = "Select * from DireccionAuxiliar Where DAuCliente = " & aIdCliente & " Order by DAuNombre"
    Set rsDA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsDA.EOF Then
        Do While Not rsDA.EOF
            cDireccion.AddItem Trim(rsDA!DAuNombre)
            cDireccion.ItemData(cDireccion.NewIndex) = rsDA!DAuDireccion
            If rsDA!DAuFactura Then lDirFact = rsDA!DAuDireccion
            rsDA.MoveNext
        Loop
        If cDireccion.ListCount > 1 Then cDireccion.BackColor = Colores.Blanco
    End If
    rsDA.Close
    
    If Val(cDireccion.Tag) = 0 And cDireccion.ListCount > 0 And lDirFact = 0 Then
        cDireccion.Text = cDireccion.List(0)
    Else
        If cDireccion.ListCount > 0 Then
            If lDirFact <> 0 Then BuscoCodigoEnCombo cDireccion, lDirFact
        End If
    End If
  
errCDA:
End Sub

Private Function loc_LoadDirFactura(ByVal lCli As Long) As Long
Dim rsDA As rdoResultset
On Error Resume Next
    
    loc_LoadDirFactura = 0
    Cons = "Select * from DireccionAuxiliar Where DAuCliente = " & lCli & " And DAuFactura = 1"
    Set rsDA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsDA.EOF Then loc_LoadDirFactura = rsDA!DAuDireccion
    rsDA.Close

End Function

'Private Sub loc_PrintContado(ByVal lDoc As Long)
'On Error GoTo ErrCrystal
'Dim Result As Integer, JobSRep1 As Integer, JobSRep2 As Integer, jobnum As Integer
'Dim NombreFormula As String, CantForm As Integer
'Dim sTexto As String
'Dim sNomMoneda As String, sSingMoneda As String, lCodMoneda As Long
'
'    lCodMoneda = Mid(lTotal.Tag, 1, InStr(1, lTotal.Tag, "|") - 1)
'    BuscoDatosMoneda lCodMoneda, sSingMoneda, sNomMoneda
'
'    Screen.MousePointer = 11
'    'Inicializo el Reporte y SubReportes
'    jobnum = crAbroReporte(gPathListados & "Contado.RPT")
'    If jobnum = 0 Then GoTo ErrCrystal
'
'    'Configuro la Impresora
'    If Trim(Printer.DeviceName) <> Trim(paIContadoN) Then SeteoImpresoraPorDefecto paIContadoN
'    If Not crSeteoImpresora(jobnum, Printer, paIContadoB) Then GoTo ErrCrystal
'
'    'Obtengo la cantidad de formulas que tiene el reporte.
'    CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
'    If CantForm = -1 Then GoTo ErrCrystal
'
'    Screen.MousePointer = 11
'
'    'Cargo Propiedades para el reporte Contado --------------------------------------------------------------------------------
'    For I = 0 To CantForm - 1
'        NombreFormula = crObtengoNombreFormula(jobnum, I)
'
'        Select Case LCase(NombreFormula)
'            Case "": GoTo ErrCrystal
'            Case "nombredocumento": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & paDContado & "'")
'            Case "cliente"
'                sTexto = ""
'                If Val(tDirNew.Tag) = 1 Then
'                    If Trim(txtCliente.Text) <> "" Then sTexto = "(" & txtCliente.Text & ")" Else sTexto = ""
'                End If
'                Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(tCliNew.Text) & " " & Trim(sTexto) & "'")
'
'            Case "direccion": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(tDirNew.Text) & "'")
'            Case "ruc":
'                If txtCliente.Cliente.Tipo = TC_Persona Then
'                    If txtCliente.Cliente.RutPersona <> "" Then Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(txtCliente.Cliente.RutPersona) & "'")
'                Else
'                    Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(txtCliente.Text) & "'")
'                End If
'            Case "codigobarras": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & CodigoDeBarras(TipoDocumento.Contado, lDoc) & "'")
'            Case "signomoneda": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(sSingMoneda) & "'")
'            Case "nombremoneda": Result = crSeteoFormula(jobnum%, NombreFormula, "'(" & sNomMoneda & ")'")
'            Case "usuario": Result = crSeteoFormula(jobnum%, NombreFormula, "'" & BuscoUsuario(miConexion.UsuarioLogueado(True), Digito:=True) & "'")
'
'            Case "textoretira"
'                Result = crSeteoFormula(jobnum%, NombreFormula, "'" & sTextRetira & "'")
'
'            Case Else: Result = 1
'        End Select
'        If Result = 0 Then GoTo ErrCrystal
'    Next
'    '--------------------------------------------------------------------------------------------------------------------------------------------
'
'    'Seteo la Query del reporte-----------------------------------------------------------------
'    Cons = "SELECT Documento.DocCodigo , Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal, Documento.DocIVA, Documento.DocVendedor" _
'            & " From " & paBD & ".dbo.Documento Documento " _
'            & " Where DocCodigo = " & lDoc
'    If crSeteoSqlQuery(jobnum%, Cons) = 0 Then GoTo ErrCrystal
'
'    'Subreporte srContado.rpt  y srContado.rpt - 01-----------------------------------------------------------------------------
'    JobSRep1 = crAbroSubreporte(jobnum, "srContado.rpt")
'    If JobSRep1 = 0 Then GoTo ErrCrystal
'
'      Cons = "SELECT Renglon.RenDocumento, Renglon.RenCantidad, Renglon.RenPrecio, Renglon.RenDescripcion, ArticuloEspecifico.AEsNombre" _
'            & " From ({ oj " & paBD & ".dbo.Renglon Renglon INNER JOIN " _
'                           & paBD & ".dbo.Articulo Articulo ON Renglon.RenArticulo = Articulo.ArtId} Left Outer Join " _
'                           & paBD & ".dbo.ArticuloEspecifico On AEsTipoDocumento = 1 And AEsDocumento = RenDocumento And AEsArticulo = RenArticulo)"
'
''    Cons = "SELECT Renglon.RenDocumento, Renglon.RenCantidad, Renglon.RenPrecio, Renglon.RenDescripcion," _
''            & " From { oj " & paBD & ".dbo.Renglon Renglon INNER JOIN " _
''                           & paBD & ".dbo.Articulo Articulo ON Renglon.RenArticulo = Articulo.ArtId}"
'    If crSeteoSqlQuery(JobSRep1, Cons) = 0 Then GoTo ErrCrystal
'
'    JobSRep2 = crAbroSubreporte(jobnum, "srContado.rpt - 01")
'    If JobSRep2 = 0 Then GoTo ErrCrystal
'    If crSeteoSqlQuery(JobSRep2, Cons) = 0 Then GoTo ErrCrystal
'    '-------------------------------------------------------------------------------------------------------------------------------------
'
'    'If crMandoAPantalla(jobnum, "Factura Contado") = 0 Then GoTo ErrCrystal
'    If crMandoAImpresora(jobnum, 1) = 0 Then GoTo ErrCrystal
'    If Not crInicioImpresion(jobnum, True, False) Then GoTo ErrCrystal
'
'    If Not crCierroSubReporte(JobSRep1) Then GoTo ErrCrystal
'    If Not crCierroSubReporte(JobSRep2) Then GoTo ErrCrystal
'
'    'crEsperoCierreReportePantalla
'
'    Screen.MousePointer = 0
'    Exit Sub
'
'ErrCrystal:
'    Screen.MousePointer = 0
'    clsGeneral.OcurrioError crMsgErr, Err.Description, "Impresión Contado"
'    On Error Resume Next
'    Screen.MousePointer = 11
'    crCierroSubReporte JobSRep1
'    crCierroSubReporte JobSRep2
'    Screen.MousePointer = 0
'    Exit Sub
'End Sub

Private Sub SeteoImpresoraPorDefecto(DeviceName As String)
Dim X As Printer

    For Each X In Printers
        If Trim(X.DeviceName) = Trim(DeviceName) Then
            Set Printer = X
            Exit For
        End If
    Next
    
End Sub

Private Function EmitirCFE(ByVal Documento As clsDocumentoCGSA, ByVal CAE As clsCAEDocumento) As String
On Error GoTo errEC
    With New clsCGSAEFactura
        .URLAFirmar = prmURLFirmaEFactura
        .TasaBasica = TasaBasica
        .TasaMinima = TasaMinima
        .ImporteConInfoDeCliente = prmImporteConInfoCliente
        Set .Connect = cBase
        If Not .GenerarEComprobante(CAE, Documento, EmpresaEmisora, paCodigoDGI) Then
            EmitirCFE = .XMLRespuesta
        End If
    End With
    Exit Function
errEC:
    EmitirCFE = "Error en firma: " & Err.Description
End Function

Public Sub CargareFacturaONOFF()
Dim rsP As rdoResultset
Dim sQy As String
    sQy = "SELECT IsNull(ParValor, 0) FROM Parametro WHERE ParNombre = 'eFacturaActiva'"
    Set rsP = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    If Not rsP.EOF Then
        prmEFacturaProductivo = rsP(0)
    End If
    rsP.Close
End Sub

Private Sub CargoDepartamentoLocalidad(ByVal idDir As Long, ByRef depto As String, ByRef loca As String)
Dim rsMV As rdoResultset
    depto = ""
    loca = ""
    Cons = "SELECT DepNombre, LocNombre " & _
        "FROM Direccion INNER JOIN Calle ON DirCalle = CalCodigo " & _
        "INNER JOIN Localidad ON CalLocalidad = LocCodigo " & _
        "INNER JOIN Departamento ON LocDepartamento = DepCodigo " & _
        "WHERE DirCodigo = " & idDir
    Set rsMV = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsMV.EOF Then
        depto = Trim(rsMV("DepNombre"))
        loca = Trim(rsMV("LocNombre"))
    End If
    rsMV.Close
End Sub

Private Sub txtCliente_BorroCliente()
    LimpioCliente
End Sub

Private Sub txtCliente_CambioTipoDocumento()
    Select Case txtCliente.DocumentoCliente
        Case DC_CI
            lRucCI.Caption = "C.I.:"
        Case DC_RUT
            lRucCI.Caption = "R.U.T.:"
        Case Else
            If txtCliente.Cliente.TipoDocumento.Nombre = "" Then
                lRucCI.Caption = "Otro:"
            Else
                lRucCI.Caption = txtCliente.Cliente.TipoDocumento.Abreviacion
            End If
            lRucCI.ForeColor = &HFF&
    End Select

End Sub

Private Sub txtCliente_Focus()
    LineHelp "Ingrese el documento del cliente."
End Sub

Private Sub txtCliente_PresionoEnter()
    If txtCliente.Cliente.Codigo > 0 And bEmitir.Enabled Then
        bEmitir.SetFocus
    End If
End Sub

Private Sub txtCliente_SeleccionoCliente()
    LimpioCliente
    SetNewCustomer txtCliente.Cliente.Codigo
    txtCliente.BuscoComentariosAlerta txtCliente.Cliente.Codigo, True
    If txtCliente.DarMsgClienteNoVender(txtCliente.Cliente.Codigo) Then
        MsgBox "Atención: NO se puede vender sin autorización. Consultar con gerencia!", vbCritical, "ATENCIÓN"
    End If
End Sub

