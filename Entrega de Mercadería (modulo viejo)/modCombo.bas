Attribute VB_Name = "modCombo"

Public gTecla As Integer        'Variable global que carga la tecla presionada en el combo
Public gSelDesde As Integer   'Variable para cargar la posición desde que se selecciona
Public gIndice As Integer        'Indice global seleccionado


'------------------------------------------------------------------------------------------------------
'   Procedimiento Selecciono: Maneja el comportamiento de textos ingresados en los
'   combos box.
'
'   Parámetros:     C:  Control combo
'                          S:  Texto del combo
'                          Tecla: Tecla que se presionó en el combo
'------------------------------------------------------------------------------------------------------

Public Sub Selecciono(C As Control, S As String, Tecla As Integer)

Dim Pos As Integer
    
    gIndice = -1
    
    If S = "" Then
        Exit Sub
    End If
    
    If Tecla = vbKeyDelete Then
        Exit Sub
    End If
    
    Pos = 0
    Do While Pos <= C.ListCount
        If UCase(Left(C.List(Pos), Len(S))) = UCase(S) Then
            C.ListIndex = Pos
            gIndice = Pos
            C.SelStart = Len(S)
            C.SelLength = Len(C.Text)
            Pos = C.ListCount
        End If
    
        Pos = Pos + 1
    Loop
    
    If Tecla = vbKeyBack Then
        C.SelStart = C.SelStart - 1
        C.SelLength = Len(C.Text)
    End If
    
    gSelDesde = C.SelStart
    
End Sub


Public Sub ComboKeyUp(C As Control)

    On Error Resume Next
    If C.ListIndex = -1 And gIndice <> -1 Then
        
        C.ListIndex = gIndice
        C.SelStart = gSelDesde
        C.SelLength = Len(C.Text)
        gSelDesde = 0
    End If
    
End Sub
