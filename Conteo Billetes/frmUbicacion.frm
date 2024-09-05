VERSION 5.00
Begin VB.Form frmUbicacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ubicación de Billetes"
   ClientHeight    =   3075
   ClientLeft      =   5205
   ClientTop       =   2295
   ClientWidth     =   3585
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUbicacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton bSalir 
      Caption         =   "&Cancelar"
      Height          =   325
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   915
   End
   Begin VB.CommandButton bGrabar 
      Caption         =   "&Aceptar"
      Height          =   325
      Left            =   2580
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   915
   End
   Begin VB.TextBox tNombre 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   60
      MaxLength       =   40
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   900
      Width           =   3435
   End
   Begin VB.CheckBox cFijo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Normalmente se mantiene"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   2100
      Width           =   2475
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6015
      TabIndex        =   4
      Top             =   0
      Width           =   6015
      Begin VB.Label lDisponibilidad 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Disponibilidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   60
         TabIndex        =   6
         Top             =   120
         Width           =   3615
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label8"
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
         Height          =   15
         Left            =   0
         TabIndex        =   5
         Top             =   480
         Width           =   6795
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Indique si en ésta ubicación, la cantidad de billetes es fija (si se mantiene de un día para el otro)."
      Height          =   555
      Left            =   60
      TabIndex        =   8
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese la ubicación de los billetes"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   660
      Width           =   3315
   End
End
Attribute VB_Name = "frmUbicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public prmIdDisponibilidad As Long
Public prmNombre As String

Public prmAddId As Long
Public prmAddTexto As String

Private Const prmIDVista = 94

Private Sub bGrabar_Click()
    
    Screen.MousePointer = 11
    On Error GoTo errorBT
    
    'Valido que no exista ubicación     -----------------------------------------------------------------------------
    Dim bAddOK As Boolean: bAddOK = True
    cons = "Select * from UbicacionBillete " & _
              " Where UBiNombre = '" & Trim(tNombre.Text) & "'" & _
              " And UBiDisponibilidad = " & prmIdDisponibilidad
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then bAddOK = False
    rsAux.Close
    '----------------------------------------------------------------------------------------------------------------------
    
    If Not bAddOK Then
        MsgBox "Ya existe una ubicación " & tNombre.Text & " para la disponibilidad " & lDisponibilidad.Caption & ".", vbInformation, "Ubicación Ingresada !!"
        Foco tNombre: Screen.MousePointer = 0
        Exit Sub
    End If
    
    cBase.BeginTrans            'COMIENZO TRANSACCION------------------------------------------     !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
    
    Dim mNewId As Long
    mNewId = 0
    cons = "Select Max(UBiCodigo) from UbicacionBillete"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux(0)) Then mNewId = rsAux(0)
    End If
    rsAux.Close
    
    mNewId = mNewId + 1
    
    'Tipo As UBiTipo, Puntaje As UBiCodigo, Texto As UBiNombre
    'Clase As UBiDisponibilidad, Valor1 As UBiFijo
    
    cons = "Select * from CodigoTexto " & _
               " Where Tipo = " & prmIDVista & " And Puntaje = " & mNewId
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    rsAux.AddNew
    rsAux!Tipo = prmIDVista
    rsAux!Puntaje = mNewId
    rsAux!Texto = Trim(tNombre.Text)
    rsAux!Clase = prmIdDisponibilidad
    If cFijo.Value = vbChecked Then rsAux!Valor1 = 1
    
    rsAux.Update: rsAux.Close
    
    cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------        !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    prmAddId = mNewId
    prmAddTexto = Trim(tNombre.Text)
    
    Unload Me
    Screen.MousePointer = 0
    Exit Sub
    
errorBT:
    clsGeneral.OcurrioError "Error al grabar los datos.", Err.Description
    Screen.MousePointer = 0: Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Error al grabar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub bSalir_Click()
    prmAddId = 0
    Unload Me
End Sub

Private Sub cFijo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bGrabar.SetFocus
End Sub

Private Sub Form_Load()

    On Error Resume Next
    prmAddId = 0: prmAddTexto = ""
    
    Me.BackColor = RGB(255, 240, 245)
    cFijo.BackColor = Me.BackColor
    tNombre.BackColor = RGB(216, 191, 216)
    bGrabar.BackColor = Me.BackColor: bSalir.BackColor = Me.BackColor
    
    InicializoForm
    
End Sub

Private Sub InicializoForm()

On Error Resume Next

    cons = "Select * from Disponibilidad Where DisID = " & prmIdDisponibilidad
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        lDisponibilidad.Caption = Trim(rsAux!DisNombre)
        lDisponibilidad.Tag = rsAux!DisID
    End If
    rsAux.Close
    
    tNombre.Text = prmNombre
    If Trim(tNombre.Text) <> "" Then
        tNombre.SelStart = Len(tNombre.Text)
        tNombre.SetFocus
    End If
    
End Sub

Private Sub tNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cFijo.SetFocus
End Sub
