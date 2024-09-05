VERSION 5.00
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Begin VB.Form MaTipo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de Productos"
   ClientHeight    =   3870
   ClientLeft      =   6015
   ClientTop       =   3810
   ClientWidth     =   7155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMaTipo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   7155
   Begin VB.TextBox tHijoDe 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   25
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5760
      TabIndex        =   20
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4440
      TabIndex        =   19
      Top             =   3360
      Width           =   1215
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   825
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   7155
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   7155
      Begin VB.Label lbHelp 
         BackStyle       =   0  'Transparent
         Caption         =   "&Hijo de:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   6855
      End
      Begin VB.Label lHelp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   18
         Top             =   420
         Width           =   45
      End
      Begin VB.Label lCaption1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipos de artículos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   60
         TabIndex        =   17
         Top             =   60
         Width           =   2295
      End
   End
   Begin VB.TextBox tBusqWeb 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   60
      TabIndex        =   9
      Top             =   2100
      Width           =   4515
   End
   Begin AACombo99.AACombo cLocReparacion 
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   1680
      Width           =   2655
      _ExtentX        =   4683
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
   Begin VB.TextBox tAbreviacion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5460
      MaxLength       =   12
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox tNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1560
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1320
      Width           =   2655
   End
   Begin AACombo99.AACombo cArancelMS 
      Height          =   315
      Left            =   1560
      TabIndex        =   11
      Top             =   2520
      Width           =   2715
      _ExtentX        =   4789
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
   Begin AACombo99.AACombo cArancelRM 
      Height          =   315
      Left            =   1560
      TabIndex        =   13
      Top             =   2880
      Width           =   2715
      _ExtentX        =   4789
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Hijo de:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lArancelRM 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   4260
      TabIndex        =   15
      Top             =   2880
      Width           =   1035
   End
   Begin VB.Label lArancelMS 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   4260
      TabIndex        =   14
      Top             =   2520
      Width           =   1035
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "&Recargo Resto:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Recargo &Mercosur:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Búsqueda &Web:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2100
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Local Reparación:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "A&breviación:"
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Nombre:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "MaTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public prmTxtNuevoTipo As String        'Texto con el nombre del nuevo Tipo (el valor prmIDTipo = -1)
Public prmTipo As Integer
Public prmHijoDe As Long

Private arrArancel() As Currency

'Propiedades.---------------------------------------
Private bTipoLlamado As Byte
Private lSeleccionado As Long

Private Function fnc_GetHijosDe(ByVal sHijosDe As String) As String
Dim rsH As rdoResultset
Dim sChild As String, sRet As String
    Set rsH = cBase.OpenResultset("Select TipCodigo From Tipo Where TipHijoDe IN (" & sHijosDe & ")", rdOpenDynamic, rdConcurValues)
    Do While Not rsH.EOF
        sChild = sChild & IIf(sChild = "", "", ", ") & rsH(0)
        rsH.MoveNext
    Loop
    rsH.Close
    If sChild <> "" Then
         sRet = fnc_GetHijosDe(sChild)
         fnc_GetHijosDe = sChild & IIf(sRet <> "", ", ", "") & sRet
    End If
End Function

Private Sub cArancelMS_Change()
    If cArancelMS.ListIndex = -1 Then
        lArancelMS.Caption = ""
    Else
        lArancelMS.Caption = Format(arrArancel(cArancelMS.ListIndex), "#,##0.000")
    End If
End Sub

Private Sub cArancelMS_Click()
    If cArancelMS.ListIndex = -1 Then
        lArancelMS.Caption = ""
    Else
        lArancelMS.Caption = Format(arrArancel(cArancelMS.ListIndex), "#,##0.000")
    End If
End Sub

Private Sub cArancelMS_GotFocus()
    lbHelp.Caption = "Seleccione el arancel"
End Sub

Private Sub cArancelMS_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then cArancelRM.SetFocus
End Sub

Private Sub cArancelMS_LostFocus()
    lbHelp.Caption = ""
End Sub

Private Sub cArancelRM_Change()
    If cArancelRM.ListIndex = -1 Then
        lArancelRM.Caption = ""
    Else
        lArancelRM.Caption = Format(arrArancel(cArancelRM.ListIndex), "#,##0.000")
    End If
End Sub

Private Sub cArancelRM_Click()
    If cArancelRM.ListIndex = -1 Then
        lArancelRM.Caption = ""
    Else
        lArancelRM.Caption = Format(arrArancel(cArancelRM.ListIndex), "#,##0.000")
    End If
End Sub

Private Sub cArancelRM_GotFocus()
    lbHelp.Caption = "Seleccione el arancel"
End Sub

Private Sub cArancelRM_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Sub cArancelRM_LostFocus()
    lbHelp.Caption = ""
End Sub



Private Sub cLocReparacion_GotFocus()
    lbHelp.Caption = "Seleccione el local de reparación para los artículos del tipo."
End Sub

Private Sub cLocReparacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tBusqWeb.SetFocus
End Sub

Private Sub cLocReparacion_LostFocus()
    lbHelp.Caption = ""
End Sub

Private Sub Command1_Click()
    AccionGrabar
End Sub

Private Sub Command2_Click()
    prmTipo = 0
    Unload Me
End Sub

Private Sub Form_Load()
Dim sAux As String

    ObtengoSeteoForm Me, 500, 500
    Me.Height = 4245
           
    Cons = "Select SucCodigo, SucAbreviacion From Sucursal Order by SucAbreviacion"
    CargoCombo Cons, cLocReparacion
    
'esto era para sacar el largo de la caract.
'    SacoLargoCampo
    
    ReDim arrArancel(0)
    Cons = "Select * from Arancel"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        cArancelMS.AddItem Trim(RsAux!AraNombre)
        cArancelMS.ItemData(cArancelMS.NewIndex) = RsAux!AraCodigo
        cArancelRM.AddItem Trim(RsAux!AraNombre)
        cArancelRM.ItemData(cArancelRM.NewIndex) = RsAux!AraCodigo
        ReDim Preserve arrArancel(cArancelRM.NewIndex)
        If Not IsNull(RsAux!AraCoeficiente) Then
            arrArancel(cArancelRM.NewIndex) = RsAux!AraCoeficiente
        Else
            arrArancel(cArancelRM.NewIndex) = 0
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If prmTipo > 0 Then
        AccionModificar
    Else
        tNombre.Text = prmTxtNuevoTipo
        If prmHijoDe > 0 Then
            Cons = "Select TipCodigo, TipNombre From Tipo Where TipCodigo = " & prmHijoDe
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                tHijoDe.Text = Trim(RsAux("TipNombre"))
                tHijoDe.Tag = RsAux("TipCodigo")
            End If
            RsAux.Close
        End If
    End If
    
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    GuardoSeteoForm Me
End Sub

Private Sub Label2_Click()
   
    tNombre.SetFocus
    tNombre.SelStart = 0
    tNombre.SelLength = Len(tNombre)

End Sub

Private Sub Label3_Click()
    With tHijoDe
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
End Sub

Private Sub Label4_Click()
    If tAbreviacion.Enabled Then tAbreviacion.SetFocus
End Sub

Private Sub Label8_Click()
    cArancelMS.SetFocus
End Sub

Private Sub Label9_Click()
    cArancelRM.SetFocus
End Sub

Sub AccionModificar()
    On Error GoTo Error
    'Cargo el RS con el pais a modificar
    Dim oTipo As New clsTipo
    If oTipo.LoadTipo(prmTipo) Then
        If oTipo.HijoDe > 0 Then
            Cons = "Select TipCodigo, TipNombre From Tipo Where TipCodigo = " & oTipo.HijoDe
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                tHijoDe.Text = Trim(RsAux("TipNombre"))
                tHijoDe.Tag = RsAux("TipCodigo")
            End If
            RsAux.Close
        End If
        tNombre.Text = oTipo.Nombre
        tAbreviacion.Text = oTipo.Abreviacion
        If oTipo.LocalReparacion > 0 Then BuscoCodigoEnCombo cLocReparacion, oTipo.LocalReparacion Else cLocReparacion.Text = ""
        tBusqWeb.Text = oTipo.BusquedaWeb
'        tArray.Text = oTipo.ArrayCaracteristicas: tArray.Tag = Trim(tArray.Text)
        If oTipo.RecargoMS > 0 Then BuscoCodigoEnCombo cArancelMS, oTipo.RecargoMS
        If oTipo.RecargoRM > 0 Then BuscoCodigoEnCombo cArancelRM, oTipo.RecargoRM
    End If
    Exit Sub
Error:
    clsGeneral.OcurrioError "Error al realizar la operación.", Err.Description
End Sub

Sub AccionGrabar()

    If Not ValidoCampos Then
        MsgBox "Los datos ingresados no son correctos o la ficha está incompleta.", vbExclamation, "ATENCIÓN"
        Exit Sub
    Else
        If Not clsGeneral.TextoValido(tNombre) Then
            MsgBox "Se ha ingresado un caracter no válido, verifique.", vbExclamation, "ATENCION"
            Exit Sub
        End If
    End If
    
    If MsgBox("¿Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    
    On Error GoTo errGrabar
    Dim oTipo As New clsTipo
    With oTipo
        .Codigo = prmTipo
        If Val(tHijoDe.Tag) > 0 Then .HijoDe = Val(tHijoDe.Tag)
        .Abreviacion = Trim(tAbreviacion.Text)
'        .ArrayCaracteristicas = Trim(tArray.Text)
        .BusquedaWeb = Trim(tBusqWeb.Text)
        .Especie = 0
        If cLocReparacion.ListIndex > -1 Then .LocalReparacion = cLocReparacion.ItemData(cLocReparacion.ListIndex)
        .Nombre = Trim(tNombre.Text)
        If cArancelMS.ListIndex > -1 Then .RecargoMS = cArancelMS.ItemData(cArancelMS.ListIndex)
        If cArancelRM.ListIndex > -1 Then .RecargoRM = cArancelRM.ItemData(cArancelRM.ListIndex)
        If .SaveTipo Then
            prmTipo = .Codigo
            prmTxtNuevoTipo = .Nombre
             Unload Me
        End If
    End With
    Set oTipo = Nothing
    Screen.MousePointer = 0
    Exit Sub
    
errGrabar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ha ocurrido un error al realizar la operación.", Err.Description
End Sub

Private Sub tAbreviacion_GotFocus()
    tAbreviacion.SelStart = 0
    tAbreviacion.SelLength = Len(tAbreviacion.Text)
    lbHelp.Caption = "Ingrese una abreviación del nombre."
End Sub

Private Sub tAbreviacion_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then cLocReparacion.SetFocus
End Sub

Private Sub tAbreviacion_LostFocus()
    lbHelp.Caption = ""
End Sub

Private Sub tBusqWeb_GotFocus()
    On Error Resume Next
    tBusqWeb.SelStart = 0
    tBusqWeb.SelLength = Len(tBusqWeb.Text)
    lbHelp.Caption = "Ingrese las palabras claves para la búsqueda web."
End Sub

Private Sub tBusqWeb_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then cArancelMS.SetFocus
End Sub

Private Sub tBusqWeb_LostFocus()
    lbHelp.Caption = ""
End Sub


Private Sub tHijoDe_Change()
    If Val(tHijoDe.Tag) > 0 Then
        tHijoDe.Tag = ""
    End If
End Sub

Private Sub tHijoDe_GotFocus()
    With tHijoDe
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    lbHelp.Caption = "Indique si el tipo será hijo de otro tipo."
End Sub

Private Sub tHijoDe_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tHijoDe.Tag) = 0 And tHijoDe.Text <> "" Then
            Cons = ""
            If prmTipo > 0 Then Cons = fnc_GetHijosDe(prmTipo)
            Cons = Cons & IIf(Cons = "", "", ", ") & prmTipo
            
            Cons = "SELECT TipCodigo, TipNombre FROM Tipo " & _
                "WHERE (TipNombre Like '" & Replace(tHijoDe.Text, " ", "%") & "%' OR TipBusqWeb Like '" & Replace(tHijoDe.Text, " ", "%") & "%')" & _
                " AND TipCodigo Not IN (" & Cons & ")"
                
            Dim oHelp As New clsListadeAyuda
            If oHelp.ActivarAyuda(cBase, Cons, 4000, 1, "Tipos de artículos") > 0 Then
                tHijoDe.Text = oHelp.RetornoDatoSeleccionado(1)
                tHijoDe.Tag = oHelp.RetornoDatoSeleccionado(0)
            End If
            Set oHelp = Nothing
            If Val(tHijoDe.Tag) > 0 Then
                With tHijoDe
                    .SelStart = 0
                    .SelLength = Len(.Text)
                End With
            End If
        Else
            tNombre.SetFocus
        End If
    End If
End Sub

Private Sub tHijoDe_LostFocus()
    lbHelp.Caption = ""
End Sub

Private Sub tNombre_GotFocus()

    tNombre.SelStart = 0
    tNombre.SelLength = Len(tNombre)
    lbHelp.Caption = "Ingrese el nombre del tipo."

End Sub

Private Sub tNombre_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn And Trim(tNombre) <> "" Then tAbreviacion.SetFocus

End Sub

Private Function ValidoCampos()

    ValidoCampos = True
    If Trim(tNombre.Text) = "" Then
        tNombre.SetFocus
        ValidoCampos = False
        Exit Function
    End If
    If Val(tHijoDe.Tag) = 0 And Trim(tHijoDe.Text) <> "" Then
        tHijoDe.SetFocus
        ValidoCampos = False
    End If
    

End Function

Private Sub tNombre_LostFocus()
    lbHelp.Caption = ""
End Sub
