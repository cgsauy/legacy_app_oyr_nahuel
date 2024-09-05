VERSION 5.00
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmTransferencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferencias entre Disponibilidades"
   ClientHeight    =   4935
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8685
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTransDisp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8685
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   9
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir del formulario"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin ComctlLib.StatusBar Status1 
         Height          =   255
         Left            =   480
         TabIndex        =   33
         Top             =   1500
         Visible         =   0   'False
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   450
         SimpleText      =   ""
         _Version        =   327682
         BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
            NumPanels       =   4
            BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "terminal"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               AutoSize        =   2
               Key             =   "usuario"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               AutoSize        =   2
               Key             =   "bd"
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               AutoSize        =   1
               Object.Width           =   8440
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox tIDMov 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      MaxLength       =   9
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox tTCEntrada 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   25
      Top             =   3060
      Width           =   975
   End
   Begin VB.TextBox tTCSalida 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   9
      Top             =   1740
      Width           =   975
   End
   Begin VB.TextBox tComentario 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   60
      TabIndex        =   31
      Top             =   4260
      Width           =   6135
   End
   Begin VB.TextBox tImporteM 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   29
      Top             =   3900
      Width           =   1695
   End
   Begin VB.TextBox tFecha 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      TabIndex        =   5
      Top             =   900
      Width           =   2055
   End
   Begin VB.TextBox tImporteE 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7080
      TabIndex        =   27
      Top             =   3060
      Width           =   1455
   End
   Begin VB.TextBox tImporteS 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7080
      TabIndex        =   11
      Top             =   1740
      Width           =   1455
   End
   Begin VB.TextBox tChVence 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5640
      MaxLength       =   11
      TabIndex        =   19
      Text            =   "10/10/2000"
      Top             =   2220
      Width           =   975
   End
   Begin VB.TextBox tChLibrado 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3960
      MaxLength       =   11
      TabIndex        =   17
      Text            =   "10/10/2000"
      Top             =   2220
      Width           =   975
   End
   Begin VB.TextBox tChNumero 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2160
      MaxLength       =   12
      TabIndex        =   15
      Text            =   "123456789012"
      Top             =   2250
      Width           =   1095
   End
   Begin VB.TextBox tChSerie 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   14
      Text            =   "AA"
      Top             =   2250
      Width           =   340
   End
   Begin VB.TextBox tChImporte 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7440
      MaxLength       =   13
      TabIndex        =   21
      Text            =   "1,000,000.00"
      Top             =   2220
      Width           =   1095
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   34
      Top             =   4680
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
      EndProperty
   End
   Begin AACombo99.AACombo cDispSalida 
      Height          =   315
      Left            =   1440
      TabIndex        =   7
      Top             =   1740
      Width           =   3135
      _ExtentX        =   5530
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
   Begin AACombo99.AACombo cChTipo 
      Height          =   315
      Left            =   240
      TabIndex        =   12
      Top             =   2220
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      ForeColor       =   0
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
   Begin AACombo99.AACombo cDisponibilidad 
      Height          =   315
      Left            =   1440
      TabIndex        =   23
      Top             =   3060
      Width           =   3015
      _ExtentX        =   5318
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
   Begin AACombo99.AACombo cTipoMovimiento 
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Top             =   900
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "&ID :"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "T.&C.:"
      Height          =   255
      Left            =   4680
      TabIndex        =   24
      Top             =   3060
      Width           =   615
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "&T.C.:"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   1740
      Width           =   615
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Importe del Movimiento en pesos"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   37
      Top             =   3540
      Width           =   8295
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Movimiento de Entrada"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   2700
      Width           =   8295
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Movimiento de Salida"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   1320
      Width           =   8295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "T&ipo:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   900
      Width           =   495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "&Comentario:"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   4260
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Im&porte:"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   3900
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "&Fecha:"
      Height          =   255
      Left            =   3780
      TabIndex        =   4
      Top             =   900
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Disponibilidad:"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   3060
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Impor&te:"
      Height          =   255
      Left            =   6360
      TabIndex        =   26
      Top             =   3060
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "I&mporte:"
      Height          =   255
      Left            =   6360
      TabIndex        =   10
      Top             =   1740
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Disponibilidad:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1740
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "&Vence:"
      Height          =   255
      Left            =   5040
      TabIndex        =   18
      Top             =   2220
      Width           =   615
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Librado:"
      Height          =   255
      Left            =   3360
      TabIndex        =   16
      Top             =   2220
      Width           =   615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "&Nº:"
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   2250
      Width           =   375
   End
   Begin VB.Label Label13 
      Caption         =   "Imp&orte:"
      Height          =   255
      Left            =   6720
      TabIndex        =   20
      Top             =   2220
      Width           =   735
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7800
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTransDisp.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTransDisp.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTransDisp.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTransDisp.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTransDisp.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTransDisp.frx":099C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuModificar 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         Shortcut        =   ^M
         Visible         =   0   'False
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu MnuSalirDelFormulario 
      Caption         =   "Salir"
      Begin VB.Menu MnuSalir 
         Caption         =   "&Del Formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cChTipo_Change()
    DeshabilitoCamposCheque DesdeTipo:=True
End Sub
Private Sub cChTipo_Click()
    DeshabilitoCamposCheque DesdeTipo:=True
End Sub
Private Sub cChTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cChTipo.ListIndex = -1 Then Exit Sub
        If cChTipo.ItemData(cChTipo.ListIndex) = 1 Then
            tChSerie.Enabled = True: tChSerie.BackColor = Obligatorio
            tChNumero.Enabled = True: tChNumero.BackColor = Obligatorio
            Foco tChSerie
        Else
            If cDisponibilidad.Enabled Then Foco cDisponibilidad Else AccionGrabar
        End If
    End If
End Sub

Private Sub DeshabilitoCamposCheque(Optional DesdeDisponibilidad As Boolean = False, Optional DesdeTipo As Boolean = False)

    If DesdeDisponibilidad Then
        cChTipo.Tag = "": cChTipo.Text = "": cChTipo.Enabled = False: cChTipo.BackColor = Inactivo
        tChNumero.Text = "": tChNumero.Enabled = False: tChNumero.BackColor = Inactivo
        tChSerie.Tag = "": tChSerie.Text = "": tChSerie.Enabled = False: tChSerie.BackColor = Inactivo
    End If
    tChNumero.Tag = 0   'ID del cheque
    
    If DesdeTipo Then
        tChSerie.Tag = "": tChSerie.Text = "": tChSerie.Enabled = False: tChSerie.BackColor = Inactivo
        tChNumero.Text = "": tChNumero.Enabled = False: tChNumero.BackColor = Inactivo
    End If
    
    tChVence.Tag = "": tChVence.Text = "": tChVence.Enabled = False: tChVence.BackColor = Inactivo
    tChLibrado.Tag = "": tChLibrado.Text = "": tChLibrado.Enabled = False: tChLibrado.BackColor = Inactivo
    tChImporte.Tag = "": tChImporte.Text = "": tChImporte.Enabled = False: tChImporte.BackColor = Inactivo

End Sub

Private Sub cDisponibilidad_GotFocus()
    With cDisponibilidad
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub cDisponibilidad_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrS
    If KeyCode = vbKeyReturn Then
        Screen.MousePointer = 11
        If cDisponibilidad.ListIndex = -1 Then
            MsgBox "No hay una disponibilidad seleccionada, verifique.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
        Cons = "Select * from Disponibilidad Where DisID = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux.EOF Then
            RsAux.Close
            Screen.MousePointer = 0
            MsgBox "No se encontro la disponibilidad seleccionada, verifique si no fue eliminada.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
        'Me guardo la moneda de la disponibilidad.
        tImporteE.Tag = RsAux!DisMoneda
        RsAux.Close
        If Val(tImporteE.Tag) <> paMonedaPesos Then
            'Busco la mayor tasa de cambio con fecha menor al primer día de este mes.
            tTCEntrada.Text = TasadeCambio(Val(tImporteE.Tag), paMonedaPesos, PrimerDia(gFechaServidor) - 1)
            tTCEntrada.Enabled = True
            'Veo si puedo sugerir si existe salida.
            If IsNumeric(tImporteS.Text) Then
                If Val(tImporteS.Tag) = Val(tImporteE.Tag) Then
                    tImporteE.Text = tImporteS.Text
                Else
                    'Si es pesos calculo en base a la tasa de cambio que tengo.
                    If Val(tImporteS.Tag) = paMonedaPesos Then
                        If IsNumeric(tTCEntrada.Text) Then
                            If CCur(tTCEntrada.Text) <> 0 Then tImporteE.Text = Format(CCur(tImporteS.Text) / CCur(tTCEntrada.Text), FormatoMonedaP)
                        End If
                    Else
                        If IsNumeric(tTCEntrada.Text) Then
                            If CCur(tTCEntrada.Text) <> 0 Then tImporteE.Text = Format((CCur(tImporteS.Text) * CCur(tTCSalida.Text)) / CCur(tTCEntrada.Text), FormatoMonedaP)
                        End If
                    End If
                End If
            End If
            Foco tTCEntrada
        Else
            tTCEntrada.Text = "1": tTCEntrada.Enabled = False
            'Sugiero si hay entrada
            If IsNumeric(tImporteS.Text) Then
                If Val(tImporteS.Tag) = Val(tImporteE.Tag) Then
                    tImporteE.Text = tImporteS.Text
                Else
                    If IsNumeric(tTCEntrada.Text) Then
                        If CCur(tTCEntrada.Text) <> 0 Then tImporteE.Text = (CCur(tImporteS.Text) * CCur(tTCSalida.Text)) / CCur(tTCEntrada.Text)
                    End If
                End If
            End If
            If IsNumeric(tImporteE.Text) Then tImporteE.Text = Format(tImporteE.Text, FormatoMonedaP)
            Foco tImporteE
        End If
        Screen.MousePointer = 0
    End If
    Exit Sub
ErrS:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error inesperado.", Err.Description
End Sub
Private Sub cDisponibilidad_LostFocus()
    cDisponibilidad.SelStart = 0
End Sub

Private Sub cDispSalida_Change()
    DeshabilitoCamposCheque DesdeDisponibilidad:=True
End Sub
Private Sub cDispSalida_Click()
    DeshabilitoCamposCheque DesdeDisponibilidad:=True
End Sub
Private Sub cDispSalida_GotFocus()
    With cDispSalida
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub cDispSalida_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrS
    If KeyCode = vbKeyReturn Then
        If cDispSalida.ListIndex = -1 Then
            MsgBox "No hay una disponibilidad seleccionada, verifique.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
        Screen.MousePointer = 11
        Cons = "Select * from Disponibilidad Where DisID = " & cDispSalida.ItemData(cDispSalida.ListIndex)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux.EOF Then
            RsAux.Close
            Screen.MousePointer = 0
            MsgBox "No se encontro la disponibilidad seleccionada, verifique si no fue eliminada.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
        'Me guardo la moneda de la disponibilidad.
        tImporteS.Tag = RsAux!DisMoneda
        If Not IsNull(RsAux!DisSucursal) Then
            cChTipo.Enabled = True: cChTipo.BackColor = Obligatorio
            cChTipo.ListIndex = 0
            tChNumero.Enabled = True: tChNumero.BackColor = Obligatorio
            tChSerie.Enabled = True: tChSerie.BackColor = Obligatorio
        End If
        RsAux.Close
        If Val(tImporteS.Tag) <> paMonedaPesos Then
            'Busco la mayor tasa de cambio con fecha menor al primer día de este mes.
            tTCSalida.Text = TasadeCambio(Val(tImporteS.Tag), paMonedaPesos, PrimerDia(gFechaServidor) - 1)
            tTCSalida.Enabled = True
            Foco tTCSalida
        Else
            tTCSalida.Text = "1": tTCSalida.Enabled = False
            Foco tImporteS
        End If
        Screen.MousePointer = 0
    End If
    Exit Sub
ErrS:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error inesperado.", Err.Description
End Sub

Private Sub cDispSalida_LostFocus()
    With cDispSalida
        .SelStart = 0
    End With
End Sub

Private Sub cTipoMovimiento_GotFocus()
    With cTipoMovimiento
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub cTipoMovimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Foco tFecha
        Case vbKeyF1
            If cTipoMovimiento.BackColor = vbWhite Then BuscarMovimientos
    End Select
End Sub

Private Sub cTipoMovimiento_LostFocus()
    With cTipoMovimiento
        .SelStart = 0
    End With
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad
    
    'Cargo los tipos de Movimientos.
    Cons = "Select TMDCodigo, TMDNombre From TipoMovDisponibilidad Where TMDTransferencia = 1 Order by TMDNombre"
    CargoCombo Cons, cTipoMovimiento
    '--------------------------------------------
    Cons = "Select DisID, DisNombre From Disponibilidad Order by DisNombre"
    CargoCombo Cons, cDisponibilidad
    CargoCombo Cons, cDispSalida
    
    cChTipo.AddItem "Cheque": cChTipo.ItemData(cChTipo.NewIndex) = 1
    cChTipo.AddItem "Orden": cChTipo.ItemData(cChTipo.NewIndex) = 2
    
    OcultoCampos
    Botones True, False, False, False, False, Toolbar1, Me
    
    FechaDelServidor
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrio un error al cargar el formulario.", Err.Description
End Sub
Private Sub Form_Unload(Cancel As Integer)
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
End Sub

Private Sub Label1_Click()
    Foco cDisponibilidad
End Sub

Private Sub Label10_Click()
    Foco tChSerie
End Sub

Private Sub Label2_Click()
    Foco tImporteE
End Sub

Private Sub Label4_Click()
    Foco cDispSalida
End Sub

Private Sub Label5_Click()
    Foco tImporteM
End Sub

Private Sub Label6_Click()
    Foco tFecha
End Sub

Private Sub Label7_Click()
    Foco cTipoMovimiento
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuEliminar_Click()
    AccionEliminar
End Sub

Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuNuevo_Click()
    AccionNuevo
End Sub

Private Sub MnuSalir_Click()
    Unload Me
End Sub

Private Sub tComentario_GotFocus()
    With tComentario
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub
Private Sub tComentario_LostFocus()
    tComentario.SelStart = 0
End Sub
Private Sub tChImporte_GotFocus()
    With tChImporte
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub tChImporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cDisponibilidad.Enabled Then
            Foco cDisponibilidad
        Else
            Foco tImporteM
        End If
    End If
End Sub

Private Sub tChImporte_LostFocus()
    With tChImporte
        .SelStart = 0
    End With
    If IsNumeric(tChImporte.Text) Then tChImporte.Text = Format(tChImporte.Text, FormatoMonedaP) Else tChImporte.Text = ""
End Sub

Private Sub tChLibrado_GotFocus()
    With tChLibrado
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub tChLibrado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tChVence
End Sub
Private Sub tChLibrado_LostFocus()
    tChLibrado.SelStart = 0
    If IsDate(tChLibrado.Text) Then tChLibrado.Text = Format(tChLibrado.Text, FormatoFP) Else tChLibrado.Text = ""
End Sub

Private Sub tChNumero_Change()
    DeshabilitoCamposCheque
End Sub

Private Sub tChNumero_GotFocus()
    With tChNumero
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tChNumero_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If cChTipo.ListIndex = -1 Then
            MsgBox "Debe seleccionar el tipo de documento para realizar la salida de la cuenta.", vbExclamation, "ATENCIÓN"
            Foco cChTipo: Exit Sub
        End If
        If tChSerie.Enabled And Trim(tChSerie.Text) = "" Then
            MsgBox "Debe ingresar el número de serie del cheque para buscarlo en la base de datos.", vbExclamation, "ATENCIÓN"
            Foco tChSerie: Exit Sub
        End If
        If Not IsNumeric(tChNumero.Text) Then
            MsgBox "Debe ingresar el número del cheque para buscarlo en la base de datos.", vbExclamation, "ATENCIÓN"
            Foco tChNumero: Exit Sub
        End If
        
        'Hay que buscar en las tablas de cheques para ver si está ingresado
        Cons = "Select * from Cheque " _
                & " Where CheIDDisponibilidad = " & cDispSalida.ItemData(cDispSalida.ListIndex) _
                & " And CheSerie = '" & Trim(tChSerie.Text) & "'" _
                & " And CheNumero = " & Trim(tChNumero.Text)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            tChNumero.Tag = RsAux!CheID
            tChImporte.Text = Format(RsAux!CheImporte, FormatoMonedaP)
            If Not IsNull(RsAux!CheLibrado) Then tChLibrado.Text = Format(RsAux!CheLibrado, "dd/mm/yyyy")
            If Not IsNull(RsAux!CheVencimiento) Then tChLibrado.Text = Format(RsAux!CheVencimiento, "dd/mm/yyyy")
        Else
            tChNumero.Tag = 0
            tChImporte.Enabled = True: tChImporte.BackColor = Obligatorio
            tChVence.Enabled = True: tChVence.BackColor = Blanco
            tChLibrado.Enabled = True: tChLibrado.BackColor = Obligatorio: tChLibrado.Text = Format(Now, "dd/mm/yyyy")
        End If
        RsAux.Close
        
        If tChLibrado.Enabled Then Foco tChLibrado: Exit Sub
        
    End If
    
End Sub

Private Sub tChNumero_LostFocus()
    tChNumero.SelStart = 0
End Sub

Private Sub tChSerie_GotFocus()
    With tChSerie
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub tChSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then Foco tChNumero
End Sub
Private Sub tChSerie_LostFocus()
    With tChSerie
        .SelStart = 0
    End With
End Sub

Private Sub tChVence_GotFocus()
    With tChVence
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tChVence_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tImporteS.Enabled Then tChImporte.Text = tImporteS.Text Else tChImporte.Text = tImporteM.Text
        Foco tChImporte
    End If
End Sub

Private Sub tChVence_LostFocus()
    With tChVence: .SelStart = 0: End With
    If IsDate(tChVence.Text) Then tChVence.Text = Format(tChVence.Text, FormatoFP) Else tChVence.Text = ""
End Sub

Private Sub tFecha_GotFocus()
    
    With tFecha
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = " Ingerse un rango ('>', '<', 'e...y..') o una fecha de búsqueda del movimiento.    [F1] Ayuda."
        
End Sub
Private Sub tFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And tFecha.BackColor = vbWhite Then
        BuscarMovimientos
    End If
End Sub
Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And tFecha.BackColor = Obligatorio Then
        If cDispSalida.Enabled Then Foco cDispSalida Else Foco cDisponibilidad
    End If
End Sub
Private Sub tFecha_LostFocus()
    With tFecha
        .SelStart = 0
    End With
    If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, FormatoFP)
    Status.SimpleText = ""
End Sub

Private Sub tIDMov_Change()
    If tIDMov.Tag <> "" Then
        OcultoCampos
        tIDMov.Tag = ""
    End If
End Sub

Private Sub tIDMov_GotFocus()
    With tIDMov
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Ingrese un id de movimiento a buscar. ([Enter] carga)"
End Sub

Private Sub tIDMov_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tIDMov.Text) Then
            OcultoCampos
            CargoMovimientoXID CLng(tIDMov.Text)
            If tIDMov.Tag = "" Then MsgBox "No existe un movimiento para el código: " & tIDMov.Text, vbInformation, "ATENCIÓN"
        Else
            cTipoMovimiento.SetFocus
        End If
    End If
End Sub

Private Sub tIDMov_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub tImporteE_GotFocus()
    With tImporteE
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tImporteE_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tImporteE.Text) And cDispSalida.ListIndex = -1 Then
            If tImporteE.Tag <> paMonedaPesos Then
                If IsNumeric(tTCEntrada.Text) Then tImporteM.Text = Format(tTCEntrada.Text * tImporteE.Text, FormatoMonedaP)
            Else
                tImporteM.Text = Format(tImporteE.Text, FormatoMonedaP)
            End If
        End If
        If IsNumeric(tImporteE.Text) Then tImporteE.Text = Format(tImporteE.Text, FormatoMonedaP): Foco tImporteM Else tImporteE.Text = ""
    End If
End Sub
Private Sub tImporteE_LostFocus()
    With tImporteE
        .SelStart = 0
    End With
    If IsNumeric(tImporteE.Text) Then tImporteE.Text = Format(tImporteE.Text, FormatoMonedaP) Else tImporteE.Text = ""
End Sub
Private Sub tImporteM_GotFocus()
    With tImporteM
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tImporteM_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tComentario
End Sub

Private Sub tImporteM_LostFocus()
    With tImporteM
        .SelStart = 0
    End With
    If IsNumeric(tImporteM.Text) Then tImporteM.Text = Format(tImporteM.Text, FormatoMonedaP) Else tImporteM.Text = ""
End Sub

Private Sub tImporteS_GotFocus()
    With tImporteS
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tImporteS_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tImporteS.Text) Then
            If tImporteS.Tag <> paMonedaPesos Then
                If IsNumeric(tTCSalida.Text) Then tImporteM.Text = Format(tImporteS.Text * tTCSalida.Text, FormatoMonedaP)
            Else
                tImporteM.Text = Format(tImporteS.Text, FormatoMonedaP)
            End If
        End If
        If cChTipo.Enabled Then
            Foco cChTipo
        Else
            If cDisponibilidad.Enabled Then
                Foco cDisponibilidad
            Else
                Foco tImporteM
            End If
        End If
    End If
End Sub

Private Sub tImporteS_LostFocus()
    With tImporteS
        .SelStart = 0
    End With
    If IsNumeric(tImporteS.Text) Then tImporteS.Text = Format(tImporteS.Text, FormatoMonedaP) Else tImporteS.Text = ""
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "nuevo"
            AccionNuevo
        Case "eliminar"
            AccionEliminar
        Case "grabar"
            AccionGrabar
        Case "cancelar"
            AccionCancelar
        Case "salir"
            Unload Me
    End Select
End Sub
Private Sub OcultoCampos()
        
    tIDMov.Enabled = True: tIDMov.BackColor = vbWindowBackground
    
    tFecha.Tag = "": tFecha.Text = "": tFecha.Enabled = True: tFecha.BackColor = vbWhite: tFecha.ForeColor = &HC00000
    cTipoMovimiento.Text = "": cTipoMovimiento.BackColor = vbWhite: cTipoMovimiento.ForeColor = &HC00000
    tImporteM.Text = "": tImporteM.Enabled = False: tImporteM.BackColor = Inactivo
    tComentario.Text = "": tComentario.Enabled = False: tComentario.BackColor = Inactivo
    
    'Campos Salida
    DeshabilitoSalida
    
    'Campos Entrada
    DeshabilitoEntrada
    
End Sub
Private Sub HabilitoCampos()
    
    tIDMov.Enabled = False: tIDMov.BackColor = vbButtonFace
    
    tFecha.Text = "": tFecha.Enabled = True: tFecha.BackColor = Obligatorio: tFecha.ForeColor = vbWindowText
    cTipoMovimiento.Text = "": cTipoMovimiento.BackColor = Obligatorio: cTipoMovimiento.ForeColor = vbWindowText
    tImporteM.Text = "": tImporteM.Enabled = True: tImporteM.BackColor = Obligatorio
    tComentario.Text = "": tComentario.Enabled = True: tComentario.BackColor = vbWhite

End Sub
Private Sub HabilitoEntrada()
    'Campos Entrada
    tImporteE.Text = "": tImporteE.Enabled = True: tImporteE.BackColor = Obligatorio
    tTCEntrada.Text = "": tTCEntrada.Enabled = True: tTCEntrada.BackColor = vbWhite
    cDisponibilidad.Text = "": cDisponibilidad.Enabled = True: cDisponibilidad.BackColor = Obligatorio
End Sub
Private Sub DeshabilitoEntrada()
    'Campos Entrada
    tImporteE.Text = "": tImporteE.Enabled = False: tImporteE.BackColor = Inactivo
    tTCEntrada.Text = "": tTCEntrada.Enabled = False: tTCEntrada.BackColor = Inactivo
    cDisponibilidad.Text = "": cDisponibilidad.Enabled = False: cDisponibilidad.BackColor = Inactivo
End Sub
Private Sub HabilitoSalida()
    'Campos Salida
    tImporteS.Text = "": tImporteS.Enabled = True: tImporteS.BackColor = Obligatorio
    tTCSalida.Text = "": tTCSalida.Enabled = True: tTCSalida.BackColor = vbWhite
    cDispSalida.Text = "": cDispSalida.Enabled = True: cDispSalida.BackColor = Obligatorio
End Sub
Private Sub DeshabilitoSalida()
    'Campos Salida
    tImporteS.Text = "": tImporteS.Enabled = False: tImporteS.BackColor = Inactivo
    tTCSalida.Text = "": tTCSalida.Enabled = False: tTCSalida.BackColor = Inactivo
    cDispSalida.Text = "": cDispSalida.Enabled = False: cDispSalida.BackColor = Inactivo
    DeshabilitoCamposCheque True, True
End Sub

Private Sub AccionNuevo()
    Screen.MousePointer = 11
    Status.SimpleText = ""
    tIDMov.Tag = "": tIDMov.Text = ""
    Botones False, False, False, True, True, Toolbar1, Me
    HabilitoCampos
    HabilitoEntrada
    HabilitoSalida
    FechaDelServidor
    tFecha.Text = Format(gFechaServidor, FormatoFP)
    Screen.MousePointer = 0
End Sub
Private Sub AccionEliminar()
Dim IDCheque As Long
Dim sOtroChMov As Boolean   'Indica si el cheque esta en otro movimiento.

    If MsgBox("¿Confirma eliminar el movimiento?", vbQuestion + vbYesNo, "ELIMINAR") = vbYes Then
        On Error GoTo ErrBCh
        Cons = "Select * From MovimientoDisponibilidadRenglon " _
            & " Where MDRIDMovimiento = " & tFecha.Tag _
            & " And MDRIDCheque > 0"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
        If Not RsAux.EOF Then IDCheque = RsAux!MDRIdCheque Else IDCheque = 0
        RsAux.Close
        sOtroChMov = False
        If IDCheque > 0 Then
            Cons = "Select * From MovimientoDisponibilidadRenglon " _
                & " Where MDRIDMovimiento <> " & tFecha.Tag _
                & " And MDRIDCheque = " & IDCheque
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
            If Not RsAux.EOF Then sOtroChMov = True
            RsAux.Close
        End If
        On Error GoTo ErrBT
        cBase.BeginTrans
        On Error GoTo ErrResumir
        Cons = "Delete MovimientoDisponibilidadRenglon Where MDRIDMovimiento = " & tFecha.Tag
        cBase.Execute (Cons)
        If IDCheque > 0 And Not sOtroChMov Then
            Cons = "Delete Cheque Where CheID = " & IDCheque
            cBase.Execute (Cons)
        End If
        Cons = "Delete MovimientoDisponibilidad Where MDiID = " & tFecha.Tag
        cBase.Execute (Cons)
        cBase.CommitTrans
        OcultoCampos
        tIDMov.Text = "": tIDMov.Tag = ""
    End If
    Screen.MousePointer = 0
    Exit Sub

ErrBCh:
    clsGeneral.OcurrioError "Ocurrio un error al buscar relación con cheques.", Err.Description
    Screen.MousePointer = 0
    Exit Sub

ErrBT:
    clsGeneral.OcurrioError "Ocurrio un error al intentar iniciar la transacción.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
ErrResumir:
    Resume ErrTrans
ErrTrans:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrio un error al intentar eliminar el movimiento.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
End Sub
Private Sub AccionCancelar()
    OcultoCampos
    Botones True, False, False, False, False, Toolbar1, Me
End Sub
Private Sub AccionGrabar()
    If ValidoDatos Then
        If MsgBox("¿Confirma almacenar los datos ingresados?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then
            GraboNuevo
        End If
    End If
End Sub
Private Sub GraboNuevo()
On Error GoTo ErrGN
Dim IDMov As Long
    Screen.MousePointer = 11
    FechaDelServidor
    cBase.BeginTrans
    On Error GoTo ErrResumir
    'Inserto eL Movimiento--------------------------------
    Cons = "Insert Into MovimientoDisponibilidad (MDiFecha, MDiHora, MDiTipo, MDiComentario) Values (" _
        & "'" & Format(tFecha.Text, sqlFormatoF) & "', '" & Format(gFechaServidor, "hh:mm:ss") & "'," & cTipoMovimiento.ItemData(cTipoMovimiento.ListIndex)
        
    If Trim(tComentario.Text) <> "" Then
        Cons = Cons & ", '" & Trim(tComentario.Text) & "')"
    Else
        Cons = Cons & ", Null)"
    End If
    cBase.Execute (Cons)
    
    Cons = "Select MAX(MDiID) From MovimientoDisponibilidad " _
        & " Where MDiFecha = '" & Format(tFecha.Text, sqlFormatoF) & "'" _
        & " And MDiHora = '" & Format(gFechaServidor, "hh:mm:ss") & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    IDMov = RsAux(0)
    RsAux.Close
    '----------------------------------------
    'Si hay Salida.
    If cDispSalida.Enabled Then
        'Cuando hay cheque tengo el ID en el Tag de tchnumero
        If cChTipo.Enabled Then
            If cChTipo.ItemData(cChTipo.ListIndex) = 1 Then
                If tChNumero.Tag = 0 Then
                    'Ingreso Cheque.
                    Cons = "Insert Into Cheque (CheIDDisponibilidad, CheSerie, CheNumero, CheImporte, CheLibrado, CheVencimiento) Values (" _
                        & cDispSalida.ItemData(cDispSalida.ListIndex) & ", '" & Trim(tChSerie.Text) & "', " & tChNumero.Text _
                        & ", " & CCur(tChImporte.Text) & ", '" & Format(tChLibrado.Text, sqlFormatoF) & "'"
                    If IsDate(tChVence.Text) Then
                        Cons = Cons & ", '" & Format(tChVence.Text, sqlFormatoF) & "')"
                    Else
                        Cons = Cons & ", Null)"
                    End If
                    cBase.Execute (Cons)
                    Cons = "Select Max(CheID) From Cheque"
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    tChNumero.Tag = RsAux(0)
                    RsAux.Close
                End If
            End If
        End If
        Cons = "Insert Into MovimientoDisponibilidadRenglon (MDRIDMovimiento, MDRIDDisponibilidad, MDRIDCheque, MDRHaber, MDRImportePesos, MDRImporteCompra) Values (" _
            & IDMov & ", " & cDispSalida.ItemData(cDispSalida.ListIndex)
        If tChNumero.Tag <> "0" Then
            Cons = Cons & "," & tChNumero.Tag
        Else
            Cons = Cons & ", 0"
        End If
        Cons = Cons & "," & CCur(tImporteS.Text) & ", " & CCur(tImporteM.Text) & ", " & CCur(tImporteS.Text) & ")"
        cBase.Execute (Cons)
    End If
    If cDisponibilidad.ListIndex > -1 Then
        Cons = "Insert Into MovimientoDisponibilidadRenglon (MDRIDMovimiento, MDRIDDisponibilidad, MDRIDCheque, MDRDebe, MDRImportePesos, MDRImporteCompra) Values (" _
            & IDMov & ", " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex) & ", 0" _
            & "," & CCur(tImporteE.Text) & ", " & CCur(tImporteM.Text) & ", " & CCur(tImporteE.Text) & ")"
        cBase.Execute (Cons)
    End If
    cBase.CommitTrans
    On Error GoTo ErrFin
    OcultoCampos
    Screen.MousePointer = 0
    On Error Resume Next
    CargoMovimientoXID IDMov
    Exit Sub
ErrGN:
    clsGeneral.OcurrioError "Ocurrio un error al intentar iniciar la transacción.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
ErrResumir:
    Resume ErrTrans
ErrTrans:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrio un error al intentar grabar los movimientos.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
ErrFin:
    clsGeneral.OcurrioError "Ocurrio un error al intentar restaurar la ficha, la información fue almacenada..", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub CargoMovimientoXID(IDMovimiento As Long)
On Error GoTo ErrCM
    Screen.MousePointer = 11
    Cons = "Select * From MovimientoDisponibilidad, MovimientoDisponibilidadRenglon " _
            & " Where MDiID = " & IDMovimiento & " And MDiID = MDRIDMovimiento"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsAux.EOF Then
        tFecha.Tag = RsAux!MDiID
        tFecha.Text = RsAux!MDiFecha
        tIDMov.Text = tFecha.Tag
        tIDMov.Tag = tFecha.Tag
        Botones True, True, True, False, False, Toolbar1, Me
        BuscoCodigoEnCombo cTipoMovimiento, RsAux!MDiTipo
        tImporteM.Text = Format(RsAux!MDRImportePesos, FormatoMonedaP)
        If Not IsNull(RsAux!MDiComentario) Then tComentario.Text = Trim(RsAux!MDiComentario)
        'Puedo tener dos renglones uno por salida y otro por entrada.
        Do While Not RsAux.EOF
            If Not IsNull(RsAux!MDRHaber) Then
                'Salida.
                BuscoCodigoEnCombo cDispSalida, RsAux!MDRIdDisponibilidad
                tImporteS.Text = Format(RsAux!MDRHaber, FormatoMonedaP)
                BuscoDatosCheque RsAux!MDRIdCheque
            Else
                'Entrada.
                BuscoCodigoEnCombo cDisponibilidad, RsAux!MDRIdDisponibilidad
                tImporteE.Text = Format(RsAux!MDRDebe, FormatoMonedaP)
            End If
            RsAux.MoveNext
        Loop
    Else
        tFecha.Tag = ""
        Botones True, False, False, False, False, Toolbar1, Me
    End If
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub
ErrCM:
    clsGeneral.OcurrioError "Ocurrio un error al intentar cargar la información del movimiento.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub BuscoDatosCheque(IDCheque As Long)
On Error GoTo ErrBDCh
Dim RsCH  As rdoResultset
    Cons = "Select * from Cheque " _
        & " Where CheID = " & IDCheque
    Set RsCH = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsCH.EOF Then
        tChSerie.Text = Trim(RsCH!CheSerie)
        tChNumero.Text = Trim(RsCH!CheNumero)
        tChNumero.Tag = IDCheque
        tChLibrado.Text = Format(RsCH!CheLibrado, FormatoFP)
        If Not IsNull(RsCH!CheVencimiento) Then tChVence.Text = Format(RsCH!CheVencimiento, FormatoFP): tChVence.Tag = tChVence.Text
        tChImporte.Text = Format(RsCH!CheImporte, FormatoMonedaP)
    End If
    RsCH.Close
    Exit Sub
ErrBDCh:
    clsGeneral.OcurrioError "Ocurrio un error al intentar cargar la información del movimiento.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Function ValidoDatos() As Boolean
    
    ValidoDatos = False
    If cTipoMovimiento.ListIndex = -1 Then
        MsgBox "Se debe ingresar un tipo de movimiento válido.", vbExclamation, "ATENCIÓN"
        Foco cTipoMovimiento: Exit Function
    End If
    If Not IsDate(tFecha.Text) Then
        MsgBox "No se ingresó una fecha válida.", vbExclamation, "ATENCIÓN"
        Foco tFecha: Exit Function
    End If
    If Not IsNumeric(tImporteM.Text) Then
        MsgBox "Se debe ingresar un importe en pesos.", vbExclamation, "ATENCIÓN"
        Foco tImporteM: Exit Function
    End If
    If CCur(tImporteM.Text) < 0 Then
        MsgBox "El importe de pago ingresado no es correcto. Verifique", vbExclamation, "ATENCIÓN"
        Foco tImporteM: Exit Function
    End If
    If Not clsGeneral.TextoValido(tComentario.Text) Then
        MsgBox "Se ingresó por lo menos una comilla simple.", vbExclamation, "ATENCIÓN"
        Foco tComentario: Exit Function
    End If
    
    'SALIDA.-----------------------------
    If cDispSalida.Enabled Then
        If cDispSalida.ListIndex = -1 Then
            MsgBox "No hay selecciondada una disponbilidad de salida válida.", vbExclamation, "ATENCIÓN"
            Foco cDispSalida: Exit Function
        End If
        If Not IsNumeric(tImporteS.Text) Then
            MsgBox "Se debe ingresar un importe válido para la salida.", vbExclamation, "ATENCIÓN"
            Foco tImporteS: Exit Function
        End If
        If tImporteS.Tag = paMonedaPesos And CCur(tImporteS.Text) <> CCur(tImporteM.Text) Then
            MsgBox "El importe de la disponibilidad no coincide con el importe del movimiento.", vbExclamation, "ATENCIÓN"
            Foco tImporteS: Exit Function
        Else
            If IsNumeric(tTCSalida.Text) Then
                If CCur(tImporteS.Text) <> Format(CCur(tImporteM.Text) / CCur(tTCSalida.Text), FormatoMonedaP) Then
                    If MsgBox("El importe ingresado como salida no coincide con el importe en pesos del movimiento." & Chr(13) & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "ATENCIÓN") = vbNo Then
                        Foco tImporteS: Exit Function
                    End If
                End If
            End If
        End If
        If cChTipo.Enabled And cChTipo.ListIndex = -1 Then
            MsgBox "Debe seleccionar el tipo de comprobante para realizar el pago.", vbExclamation, "ATENCIÓN"
            Foco cChTipo: Exit Function
        End If
        If tChSerie.Enabled And Trim(tChSerie.Text) = "" Then
            MsgBox "Debe ingresar la serie del comprobante de pago.", vbExclamation, "ATENCIÓN"
            Foco tChSerie: Exit Function
        End If
        If tChNumero.Enabled And Not IsNumeric(tChNumero.Text) Then
            MsgBox "Debe ingresar el número del comprobante de pago.", vbExclamation, "ATENCIÓN"
            Foco tChNumero: Exit Function
        End If
        If tChLibrado.Enabled And Not IsDate(tChLibrado.Text) Then
            MsgBox "Debe ingresar la fecha de librado del comprobante de pago.", vbExclamation, "ATENCIÓN"
            Foco tChLibrado: Exit Function
        End If
        If tChVence.Enabled And Not IsDate(tChVence.Text) And (tChVence.Text) <> "" Then
            MsgBox "La fecha de vencimiento del comprobante de pago no es correcta.", vbExclamation, "ATENCIÓN"
            Foco tChVence: Exit Function
        End If
        If tChImporte.Enabled And Not IsNumeric(tChImporte.Text) Then
            MsgBox "Debe ingresar el importe total del comprobante de pago.", vbExclamation, "ATENCIÓN"
            Foco tChImporte: Exit Function
        End If
    End If
    
    'ENTRADA.-----------------------------
    If cDisponibilidad.Enabled Then
        If cDisponibilidad.ListIndex = -1 Then
            MsgBox "No hay selecciondada una disponbilidad de entrada válida.", vbExclamation, "ATENCIÓN"
            Foco cDisponibilidad: Exit Function
        End If
        If Not IsNumeric(tImporteE.Text) Then
            MsgBox "Se debe ingresar un importe válido para la entrada.", vbExclamation, "ATENCIÓN"
            Foco tImporteE: Exit Function
        End If
        If tImporteE.Tag = paMonedaPesos And CCur(tImporteE.Text) <> CCur(tImporteM.Text) Then
            MsgBox "El importe de la disponibilidad no coincide con el importe del movimiento.", vbExclamation, "ATENCIÓN"
            Foco tImporteE: Exit Function
        Else
            If IsNumeric(tTCEntrada.Text) Then
                If CCur(tImporteE.Text) <> CCur(Format(CCur(tImporteM.Text) / CCur(tTCEntrada.Text), FormatoMonedaP)) Then
                    If MsgBox("El importe ingresado como entrada no coincide con el importe en pesos del movimiento." & Chr(13) & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "ATENCIÓN") = vbNo Then
                        Foco tImporteE: Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    ValidoDatos = True
End Function

Private Sub tTCEntrada_GotFocus()
    With tTCEntrada
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tTCEntrada_KeyPress(KeyAscii As Integer)
    If vbKeyReturn = KeyAscii Then Foco tImporteE
End Sub

Private Sub tTCEntrada_LostFocus()
    With tTCEntrada
        .SelStart = 0:
    End With
    If IsNumeric(tImporteS.Text) Then
        If Val(tImporteS.Tag) = Val(tImporteE.Tag) Then
            tImporteE.Text = tImporteS.Text
        Else
            'Si es pesos calculo en base a la tasa de cambio que tengo.
            If Val(tImporteS.Tag) = paMonedaPesos Then
                If IsNumeric(tTCEntrada.Text) Then
                    If CCur(tTCEntrada.Text) <> 0 Then tImporteE.Text = Format(CCur(tImporteS.Text) / CCur(tTCEntrada.Text), FormatoMonedaP)
                End If
            Else
                If IsNumeric(tTCEntrada.Text) Then
                    If CCur(tTCEntrada.Text) <> 0 Then tImporteE.Text = Format((CCur(tImporteS.Text) * CCur(tTCSalida.Text)) / CCur(tTCEntrada.Text), FormatoMonedaP)
                End If
            End If
        End If
    End If
End Sub
Private Sub tTCSalida_GotFocus()
    With tTCSalida
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub tTCSalida_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tImporteS
End Sub
Private Sub tTCSalida_LostFocus()
    With tTCSalida
        .SelStart = 0:
    End With
End Sub
Private Sub BuscarMovimientos()
On Error GoTo ErrBM

    If Trim(tFecha.Text) <> "" And ValidoPeriodoFechas(tFecha.Text) = "" Then
        MsgBox "El formato de fecha ingresado no es válido.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    Cons = "Select 'ID Movimiento' = MDiID, Fecha = MDiFecha, Tipo = RTrim(TMDNombre), Salida = Sal.DisNombre, Entrada = Ent.DisNombre, Debe = MDRDebe, Haber = MDRHaber" _
        & " From MovimientoDisponibilidad, TipoMovDisponibilidad, MovimientoDisponibilidadRenglon " _
            & "Left Outer Join Disponibilidad Sal On Sal.DisID = MDRIDDisponibilidad And MDRHaber IS Not Null " _
            & "Left Outer Join Disponibilidad Ent On Ent.DisID = MDRIDDisponibilidad And MDRDebe IS Not Null " _
        & " Where MDiID = MDRIDMovimiento "
    
    If cTipoMovimiento.ListIndex > -1 Then Cons = Cons & " And MDiTipo = " & cTipoMovimiento.ItemData(cTipoMovimiento.ListIndex)
    If Trim(tFecha.Text) <> "" Then Cons = Cons & ConsultaDeFecha("And", "MDiFecha", tFecha.Text)
    Cons = Cons & " And MDiTipo = TMDCodigo"
    
    Cons = Cons & " Order by MDiFecha Desc, MDiHora Desc"
    
    OcultoCampos
    
    Dim LiAyuda As New clsListadeAyuda
    LiAyuda.ActivoListaAyudaSQL Cons, miConexion.TextoConexion("Comercio")
    If LiAyuda.ItemSeleccionadoSQL <> "" Then OcultoCampos: CargoMovimientoXID LiAyuda.ItemSeleccionadoSQL
    Set LiAyuda = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrBM:
    clsGeneral.OcurrioError "Ocurrio un error al buscar los movimientos.", Err.Description
    Screen.MousePointer = 0
End Sub
