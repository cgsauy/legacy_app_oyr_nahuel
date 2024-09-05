VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form MaIncoterm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Incoterms"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   3390
   Icon            =   "MaIncoterm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   3390
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ListView lLista 
      Height          =   1575
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   700
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox tNombre 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1080
      MaxLength       =   7
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   3180
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
      EndProperty
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
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Incoterms &Ingresados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Nombre"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaIncoterm.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaIncoterm.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaIncoterm.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaIncoterm.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaIncoterm.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaIncoterm.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaIncoterm.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MaIncoterm.frx":0DC8
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
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "&Del formulario   Alt+F4"
      End
   End
End
Attribute VB_Name = "MaIncoterm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RsInco As rdoResultset
Private sNuevo As Boolean, sModificar As Boolean

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    sNuevo = False: sModificar = False
    tNombre.Enabled = False: tNombre.BackColor = Inactivo
    lLista.Enabled = True
    CargoLista
    Exit Sub
ErrLoad:
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    RsInco.Close
    
End Sub

Private Sub Label1_Click()

    If tNombre.Enabled Then
        tNombre.SetFocus
        tNombre.SelStart = 0
        tNombre.SelLength = Len(tNombre.Text)
    End If
        
End Sub

Private Sub lLista_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = "Lista de Incoterms ingresados."
    
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

Private Sub MnuModificar_Click()

    AccionModificar

End Sub

Private Sub MnuNuevo_Click()

    AccionNuevo

End Sub

Private Sub MnuVolver_Click()

    Unload Me

End Sub

Sub AccionNuevo()
    
    'Prendo Señal que es uno nuevo.
    sNuevo = True
    
    Call Botones(False, False, False, True, True, Toolbar1, Me)
        
    tNombre.Enabled = True
    tNombre.BackColor = Obligatorio
    lLista.Enabled = False

End Sub

Sub AccionModificar()

    sModificar = True
    
    Call Botones(False, False, False, True, True, Toolbar1, Me)
    tNombre.Enabled = True
    tNombre.BackColor = Obligatorio
    lLista.Enabled = False
    
    On Error GoTo Error
    'Cargo el RS con el continente a modificar
    RsInco.Close
    Cons = "Select * from Incoterm Where IncCodigo = " & Right(lLista.SelectedItem.Key, Len(lLista.SelectedItem.Key) - 1)
    Set RsInco = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsInco.EOF Then
        tNombre.Text = Trim(RsInco!IncNombre)
        
    Else
        sModificar = False
        MsgBox "El registro seleccionado ha sido eliminado", vbInformation, "ATENCIÓN"
        RsInco.Close
        CargoLista
        tNombre.Enabled = False
        tNombre.BackColor = Inactivo
        lLista.Enabled = True
    End If
    Exit Sub
    
Error:
    MsgBox "Ha ocurrido un error al realizar la operación.", vbCritical, "ERROR"
    sModificar = False
    CargoLista
    tNombre.Enabled = False
    tNombre.BackColor = Inactivo
    lLista.Enabled = True

End Sub

Sub AccionGrabar()

    'Se controla que se ingrese un nombre.
    If Trim(tNombre) = "" Then
            MsgBox "Los datos ingresados no son correctos o la ficha está incompleta.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    If MsgBox("¿Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then
        If sNuevo Then                  'Nuevo----------
            On Error GoTo errGrabar
            RsInco.AddNew
            RsInco!IncNombre = Trim(tNombre.Text)
            RsInco.Update
            sNuevo = False
   
        Else                                    'Modificar----
        
            On Error GoTo errGrabar
            RsInco.Edit
            RsInco!IncNombre = Trim(tNombre.Text)
            RsInco.Update
            sModificar = False
            
        End If
        RsInco.Close
        CargoLista
        tNombre.Text = ""
        tNombre.Enabled = False
        tNombre.BackColor = Inactivo
        lLista.Enabled = True

    End If
    Exit Sub
    
errGrabar:
    MsgBox "Ha ocurrido un error al realizar la operación." & MapeoError, vbCritical, "ERROR"
    RsInco.Requery
    

End Sub

Sub AccionEliminar()

    'Verifico para eliminar
    'Cons = "Select * from Pais Where PaiContinente = " & Right(lLista.SelectedItem.Key, Len(lLista.SelectedItem.Key) - 1)
    'Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    'If Not RsAux.EOF Then
    '    MsgBox "No es posible eliminar el registro seleccionado. Verifique las relaciones de datos.", vbInformation, "ATENCIÓN"
    '    RsAux.Close
    '    Exit Sub
    'End If
    'RsAux.Close
    
    If MsgBox("¿Confirma eliminar el incoterm: '" & lLista.SelectedItem & "'?", vbQuestion + vbYesNo, "ELIMINAR") = vbYes Then
        
        On Error GoTo Error
        'Cargo el RS con el registro a eliminar
        RsInco.Close
        Cons = "Select * from Incoterm Where IncCodigo = " & Right(lLista.SelectedItem.Key, Len(lLista.SelectedItem.Key) - 1)
        Set RsInco = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If Not RsInco.EOF Then
            RsInco.Delete
        Else
            MsgBox "El registro seleccionado ha sido eliminado por otra terminal", vbInformation, "ATENCIÓN"
        End If
        
        RsInco.Close
        CargoLista
        
        tNombre.Text = ""
        tNombre.Enabled = False
        tNombre.BackColor = Inactivo
        lLista.Enabled = True
    End If
    Exit Sub
    
Error:
    MsgBox "Ha ocurrido un error al realizar la operación." & MapeoError, vbCritical, "ERROR"
    RsInco.Requery

End Sub

Sub AccionCancelar()

    sNuevo = False
    sModificar = False
    
    tNombre.Text = ""
    tNombre.Enabled = False
    tNombre.BackColor = Inactivo
    
    lLista.Enabled = True
    lLista.SetFocus
    
    If lLista.ListItems.Count > 0 Then
        Call Botones(True, True, True, False, False, Toolbar1, Me)
    Else
        Call Botones(True, False, False, False, False, Toolbar1, Me)
    End If
        
End Sub

Private Sub tNombre_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And Trim(tNombre.Text) <> "" Then
        AccionGrabar
    End If
    
End Sub

Private Sub tNombre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Status.SimpleText = "Ingrese un nombre para el Incoterm."
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        
        Case "nuevo"
            AccionNuevo
        
        Case "modificar"
            AccionModificar
        
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

'-----------------------------------------------------------------------------------------------
'   Carga la lista con los datos de la BD.
'-----------------------------------------------------------------------------------------------
Private Sub CargoLista()

    lLista.ListItems.Clear
    Cons = "Select * from Incoterm"
    Set RsInco = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsInco.EOF
    
        Set itmx = lLista.ListItems.Add(, "A" + Str(RsInco!IncCodigo), Trim(RsInco!IncNombre))
        RsInco.MoveNext
        
    Loop
    
    If lLista.ListItems.Count > 0 Then
        Call Botones(True, True, True, False, False, Toolbar1, Me)
    Else
        Call Botones(True, False, False, False, False, Toolbar1, Me)
    End If
    
End Sub

