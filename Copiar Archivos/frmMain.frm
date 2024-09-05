VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copiar Archivos"
   ClientHeight    =   2790
   ClientLeft      =   3525
   ClientTop       =   3885
   ClientWidth     =   4380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4380
   Begin VB.ComboBox cDestino 
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   960
      Width           =   4215
   End
   Begin VB.ComboBox cOrigen 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   300
      Width           =   4215
   End
   Begin VB.CommandButton bCopy 
      Caption         =   "Copiar"
      Height          =   315
      Left            =   3360
      TabIndex        =   5
      Top             =   1380
      Width           =   915
   End
   Begin VB.CheckBox chSub 
      Caption         =   "Incluir Subcarpetas"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   1380
      Width           =   1935
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   15
      Left            =   0
      TabIndex        =   7
      Top             =   1740
      Width           =   5115
   End
   Begin VB.Label lCopy 
      Caption         =   "Copiando ..."
      Height          =   795
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   4155
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Destino:"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Origen:"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFolders() As String

Dim I As Integer
Dim aQTotal As Long, aQCopy As Long, aQError As Long
Dim bUnload As Boolean

Private Sub bCopy_Click()

    On Error GoTo errCopy
    
    If Trim(cOrigen.Text) = "" Or Trim(cDestino.Text) = "" Then Exit Sub
    
    fnc_ArmoArchivo
  
    Copiando True
    
    'Valido si las carpetas existen ------------------------------------------------------------------------
    Dim aTxt As String, fileDateD As Date
    If Right(Trim(cOrigen.Text), 1) <> "\" Then cOrigen.Text = Trim(cOrigen.Text) & "\"
    If Right(Trim(cDestino.Text), 1) <> "\" Then cDestino.Text = Trim(cDestino.Text) & "\"
    
    aTxt = Trim(cOrigen.Text)
    If Not ExisteFolder(aTxt) Then
        MsgBox "La carpeta 'Origen' no existe." & vbCrLf & "Verifique si tiene permisos de acceso.", vbExclamation, "Carpeta Inexistente"
        Copiando False
        Exit Sub
    End If
    
    aTxt = Trim(cDestino.Text)
    If Not ExisteFolder(aTxt) Then
        MsgBox "La carpeta 'Destino' no existe." & vbCrLf & "Verifique si tiene permisos de acceso.", vbExclamation, "Carpeta Inexistente"
        Copiando False
        Exit Sub
    End If
    '----------------------------------------------------------------------------------------------------------
    
    Screen.MousePointer = 11
    aQTotal = 0: aQCopy = 0: aQError = 0
    If Not ArmoDirectorios Then
        Copiando False
        Exit Sub
    End If
        
    For I = LBound(arrFolders) To UBound(arrFolders)
        CopioArchivos arrFolders(I)
    Next
    
    lCopy.Caption = aQTotal & " archivos procesados. " & vbCrLf & _
                            aQCopy & " archivos actualizados" & vbCrLf & _
                            aQError & " errores"
                            
    lCopy.Refresh
    
    Copiando False
    Screen.MousePointer = 0
    If bUnload Then Unload Me
    Exit Sub
    
errCopy:
    clsGeneral.OcurrioError "Error al copiar los archivos.", Err.Description
    lCopy.Caption = aQTotal & " archivos procesados. " & vbCrLf & _
                            aQCopy & " archivos actualizados"
    lCopy.Refresh
    Copiando False
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim sParam As String
    
    bUnload = False
    lCopy.Caption = ""
    cOrigen.Text = ""
    cDestino.Text = ""
    
    fnc_CargoCombos
    
    sParam = Trim(Command())
    'sParam = "C:\Query\zTest\|C:\Query\Borrame\|0"
    If sParam <> "" Then
        Dim arrParam() As String
        arrParam = Split(sParam, "|")
        cOrigen.Text = Trim(arrParam(0))
        cDestino.Text = Trim(arrParam(1))
        If Val(arrParam(2)) = 0 Then chSub.Value = vbUnchecked Else chSub.Value = vbChecked
    
        Me.Show
        Me.Refresh
        bUnload = True
        Call bCopy_Click
    End If
    
    
End Sub

Private Function ArmoDirectorios() As Boolean

Dim aDe As Integer, aHasta As Integer
Dim bSalir As Boolean
Dim aFolder As String
    
    ArmoDirectorios = False
    
    ReDim Preserve arrFolders(0)
    arrFolders(0) = Trim(cOrigen.Text)
    
    If chSub.Value = vbChecked Then
        aHasta = -1
        Do While Not bSalir
            aDe = aHasta + 1
            aHasta = UBound(arrFolders)
            
            For I = aDe To aHasta
                aFolder = Trim(arrFolders(I))
                If Not CargoFolders(aFolder) Then Exit Function
            Next I
        
            If UBound(arrFolders) = aHasta Then bSalir = True
        Loop
    End If
    
    ArmoDirectorios = True
    
End Function

Private Function CargoFolders(myPath As String) As Boolean
On Error GoTo errCF
CargoFolders = False

Dim myName As String
Dim idx As Integer
            
    myName = Dir(myPath, vbDirectory)   ' Retrieve the first entry.
    Do While myName <> ""   ' Start the loop.
       ' Ignore the current directory and the encompassing directory.
       If myName <> "." And myName <> ".." Then
          ' Use bitwise comparison to make sure MyName is a directory.
          If (GetAttr(myPath & myName) And vbDirectory) = vbDirectory Then
                idx = UBound(arrFolders) + 1
                ReDim Preserve arrFolders(idx)
                arrFolders(idx) = myPath & myName & "\"
                                
                            
          End If   ' it represents a directory.
       End If
       myName = Dir   ' Get next entry.
    Loop
    
    CargoFolders = True
    Exit Function
    
errCF:
    clsGeneral.OcurrioError "Error al procesar el directorio.", Err.Description
End Function


Private Function CopioArchivos(sCarpeta As String)

Dim sDestino As String
    
Dim Archivo, C, fileDateO As Date, fileDateD As Date

    If Trim(LCase(sCarpeta)) = Trim(LCase(cOrigen.Text)) Then
        sDestino = Trim(cDestino.Text)
    Else
        sDestino = Trim(cDestino.Text) & Mid(sCarpeta, Len(cOrigen.Text) + 1)
        'Veo si existe carpeta destino
        fileDateD = ArchivoDestino(Mid(sDestino, 1, Len(sDestino) - 1))
        If fileDateD = CDate("01/01/1900") Then
            MkDir (sDestino)
        End If
    End If
    
    Archivo = Dir(sCarpeta, 1)
    
    Do While Trim(Archivo) <> ""
        aQTotal = aQTotal + 1
        fileDateO = FileDateTime(sCarpeta & Archivo)
        fileDateD = ArchivoDestino(sDestino & Archivo)
        
        If fileDateO > fileDateD Then
            lCopy.Caption = "Copiando ..." & vbCrLf & sCarpeta & Archivo & "    A ..." & vbCrLf & sDestino & Archivo
            lCopy.Refresh
            If myFileCopy(sCarpeta & Archivo, sDestino & Archivo) Then
                aQCopy = aQCopy + 1
            Else
                aQError = aQError + 1
            End If
        End If
        Archivo = Dir ' Obtiene siguiente entrada.
    Loop

End Function

Private Function myFileCopy(mOrigen As String, mDestino As String) As Boolean
    
    On Error GoTo errCpy
    myFileCopy = True
    FileCopy Trim(mOrigen), Trim(mDestino)
    Exit Function
    
errCpy:
    myFileCopy = False
End Function
Private Function ArchivoDestino(myFile As String) As Date
    On Error Resume Next
    ArchivoDestino = "01/01/1900"
    ArchivoDestino = FileDateTime(myFile)
    
End Function

Private Function ExisteFolder(myFile As String) As Boolean
    On Error GoTo errExiste
    ExisteFolder = False
    If Dir(myFile, vbDirectory) <> "" Then ExisteFolder = True
    Exit Function

errExiste:
End Function

Private Sub Copiando(bEstado As Boolean)
    
    bCopy.Enabled = Not bEstado
    cOrigen.Enabled = Not bEstado
    cDestino.Enabled = Not bEstado
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Function fnc_ArmoArchivo()
On Error GoTo errFnc
Dim idx As Integer, mSTR As String
Dim bSAVE As Boolean, bEsta As Boolean

    If cOrigen.ListIndex <> -1 And cDestino.ListIndex <> -1 Then Exit Function
    mSTR = ""
    
    If cOrigen.ListIndex = -1 Then
        bEsta = False
        '1) Valido si el texto Ya esta en el combo      -----------------------------------
        For idx = 0 To cOrigen.ListCount - 1
            If LCase(Trim(cOrigen.Text)) = Trim(LCase(cOrigen.List(idx))) Then
                bEsta = True
                Exit For
            End If
        Next
        If Not bEsta Then cOrigen.AddItem cOrigen.Text
    End If
    
    mSTR = mSTR & "[ORIGEN]" & vbCrLf
        
    For idx = 0 To cOrigen.ListCount - 1
        mSTR = mSTR & Trim(cOrigen.List(idx)) & vbCrLf
    Next

    If cDestino.ListIndex = -1 Then
        bEsta = False
        '1) Valido si el texto Ya esta en el combo      -----------------------------------
        For idx = 0 To cDestino.ListCount - 1
            If LCase(Trim(cDestino.Text)) = Trim(LCase(cDestino.List(idx))) Then
                bEsta = True
                Exit For
            End If
        Next
        If Not bEsta Then cDestino.AddItem cDestino.Text
    End If
    
    mSTR = mSTR & "[DESTINO]" & vbCrLf
        
    For idx = 0 To cDestino.ListCount - 1
        mSTR = mSTR & Trim(cDestino.List(idx)) & vbCrLf
    Next
    
   fnc_GrabarArchivo mSTR
   Exit Function
   
errFnc:
    clsGeneral.OcurrioError "Error al grabar el archivo.", Err.Description
End Function

Private Function fnc_CargoCombos()
On Error GoTo errCargar
Dim mSTR As String
Dim idx As Integer
    mSTR = fnc_LeoArchivo
    
    If Trim(mSTR) <> "" Then
        Dim mVAL() As String, mPaso As Integer
        
        mVAL = Split(mSTR, vbCrLf)
        mPaso = 1
        
        For idx = LBound(mVAL) To UBound(mVAL)
            If Trim(mVAL(idx)) = "[ORIGEN]" Then
                mPaso = 1
            ElseIf Trim(mVAL(idx)) = "[DESTINO]" Then
                mPaso = 2
            Else
                If Trim(mVAL(idx)) <> "" Then
                    If mPaso = 1 Then cOrigen.AddItem mVAL(idx)
                    If mPaso = 2 Then cDestino.AddItem mVAL(idx)
                End If
            End If
        Next
        
    End If
    Exit Function
    
errCargar:
    clsGeneral.OcurrioError "Error al leer los datos del archivo.", Err.Description
End Function
