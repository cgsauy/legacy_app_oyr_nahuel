VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type tMargen
    Left As Single
    Right As Single
    Top As Single
    Bottom As Single
End Type
Private Type tDevice
    Nombre As String
    Bandeja As Integer
    Papel As Integer
    Margenes As tMargen
End Type
Private oDevice As tDevice
    
Private oFieldPrint As clsFieldsPrint

Public Sub SetDevice(ByVal sDevice As String, iBandeja As Integer, ByVal iPapel As Integer)
    With oDevice
        .Papel = iPapel
        .Bandeja = iBandeja
        .Nombre = sDevice
    End With
End Sub

Public Function LoadFileData(ByVal sFile As String) As Boolean
On Error GoTo errLFD
Dim sData As String, sNom As String, arrLine() As String
Dim oFile As New clsorCGSA
Dim iQ As Integer
Dim iTipo As eTipoField

    Screen.MousePointer = 11
    If oFile.GetDataFile2(sFile, sData) = 0 Then
        
        Dim oPropField As clsPropFieldPrint
        Dim oXY As clsPosXY
        
        arrLine = Split(sData, vbCrLf)
        'Si tengo algo y me llaman de nuevo.
        Set oFieldPrint = Nothing
        Set oFieldPrint = New clsFieldsPrint
        For iQ = 0 To UBound(arrLine)
            sData = Trim(arrLine(iQ))
            sNom = ""
            If InStr(1, sData, "[") = 1 And InStr(1, sData, "]") > 1 Then
                If InStr(1, sData, "[mleft]", vbTextCompare) = 1 Then
                    sData = Replace(sData, "[mleft]", "", , , vbTextCompare)
                    If IsNumeric(sData) Then oDevice.Margenes.Left = sData
                ElseIf InStr(1, sData, "[mtop]", vbTextCompare) = 1 Then
                    sData = Replace(sData, "[mtop]", "", , , vbTextCompare)
                    If IsNumeric(sData) Then oDevice.Margenes.Top = sData
                ElseIf InStr(1, sData, "[mright]", vbTextCompare) = 1 Then
                    sData = Replace(sData, "[mright]", "", , , vbTextCompare)
                    If IsNumeric(sData) Then oDevice.Margenes.Right = sData
                ElseIf InStr(1, sData, "[mbottom]", vbTextCompare) = 1 Then
                    sData = Replace(sData, "[mbottom]", "", , , vbTextCompare)
                    If IsNumeric(sData) Then oDevice.Margenes.Bottom = sData
                Else
                    Select Case LCase(Mid(sData, 1, 7))
                        Case "[label."
                            sNom = Mid(sData, 1, InStr(1, sData, "]"))
                            sNom = Trim(Replace(Replace(sNom, "[label.", ""), "]", ""))
                            sData = Mid(sData, InStr(1, sData, "]") + 1)
                            iTipo = Label
                        Case "[tabla."
                            sNom = Mid(sData, 1, InStr(1, sData, "]"))
                            sNom = Trim(Replace(Replace(sNom, "[label.", ""), "]", ""))
                            sData = Mid(sData, InStr(1, sData, "]") + 1)
                            iTipo = Tabla
                        Case Else
                            sNom = Mid(sData, 2, InStr(1, sData, "]") - 1)
                            sNom = Trim(Replace(sNom, "]", ""))
                            sData = Mid(sData, InStr(1, sData, "]") + 1)
                            iTipo = Campo
                    End Select
                    
                    Dim arrPos() As String
                    Dim arrRegXY() As String, arrXY() As String
                    Dim iR As Byte
                    
                    If sData <> "" Then
                        arrPos = Split(sData, ":")
                        If UBound(arrPos) > 0 And arrPos(0) <> "" Then
                            'Posición:
                            '   0 tengo el largo y en tabla el largo de c/col.
                            '   1 tengo el align y en tabla el nombre de todos los campos
                            '   2 si escribo en bold
                            '   3 tamaño de la fuente
                            '   4 string con la pos x e y para c/vez que se repita.
                            Set oPropField = New clsPropFieldPrint
                            With oPropField
                                If iTipo <> Tabla Then
                                    .Width = IIf(IsNumeric(arrPos(0)), arrPos(0), -1)
                                    .Align = IIf(IsNumeric(arrPos(1)), arrPos(1), taLeftTop)
                                Else
                                    .Width = arrPos(0)
                                    .Align = 0
                                    .TablaCampos = arrPos(1)
                                End If
                                If IsNumeric(arrPos(2)) Then .Bold = (Val(arrPos(2)) = 1)
                                If IsNumeric(arrPos(3)) Then .FontSize = arrPos(3) Else .FontSize = 9
                                
                                
                                'Armo posición
                                If UBound(arrPos) >= 4 Then
                                    arrRegXY = Split(arrPos(4), "|")
                                    
                                    
                                    For iR = 0 To UBound(arrRegXY)
                                        If InStr(1, arrRegXY(iR), ";") > 0 Then
                                            arrXY = Split(arrRegXY(iR), ";")
                                            
                                            Set oXY = New clsPosXY
                                            'Pos 0 = x
                                            'Pos 1 = y
                                            If IsNumeric(arrXY(0)) Then oXY.posX = arrXY(0)
                                            If IsNumeric(arrXY(1)) Then oXY.posY = arrXY(1)
                                            
                                    'Agrego las posiciones.
                                            .Posicion.AddPosicion oXY
                                        End If
                                    Next
                                End If
                                If UBound(arrPos) >= 5 Then
                                    .FontName = arrPos(5)
                                End If
                            End With
                        End If
                        If sNom <> "" Then
                            Dim oFieldP As New clsRegFieldPrint
                            With oFieldP
                                .Nombre = sNom
                                .Tipo = iTipo
                                Set .PropsFP = oPropField
                            End With
                        
                        'agrego el objeto a la colección.
                           oFieldPrint.AddField oFieldP
                           Set oFieldP = Nothing
                        End If
                    End If
                End If
            End If
        Next iQ
        LoadFileData = True
    Else
        MsgBox "Error al leer el archivo con la información de impresión." & vbCrLf & vbCrLf & "Detalle: " & sData, vbExclamation, "Error"
    End If
    Set oFile = Nothing
    Screen.MousePointer = 0
    Exit Function
errLFD:
    MsgBox Err.Description
    Screen.MousePointer = 0
End Function

Private Function fnc_InitPrinter(ByRef vsView As vsPrinter) As Boolean
    With vsView
        .StartDoc
        If .Error <> 0 Then
            .EndDoc
            MsgBox "Error al iniciar el objeto de impresión." & vbCrLf & "Error: " & Err.Description, vbCritical, "ATENCIÓN"
        Else
            .Device = oDevice.Nombre
            On Error Resume Next
            .PaperSize = oDevice.Papel
            .PaperBin = oDevice.Bandeja
            
            If oDevice.Margenes.Bottom > 0 Then .MarginBottom = oDevice.Margenes.Bottom
            If oDevice.Margenes.Left > 0 Then .MarginLeft = oDevice.Margenes.Left
            If oDevice.Margenes.Top > 0 Then .MarginTop = oDevice.Margenes.Top
            If oDevice.Margenes.Right > 0 Then .MarginRight = oDevice.Margenes.Right
            
            fnc_InitPrinter = True
        End If
    End With
End Function

Private Sub loc_PrintFieldLabel(ByVal oFP As clsRegFieldPrint, ByVal vsView As vsPrinter, ByVal vValue As Variant)
    Dim iP As Integer
    Select Case oFP.PropsFP.Align
        Case 0: vsView.TextAlign = taLeftTop
        Case taCenterTop: vsView.Text = taCenterTop
        Case taRightTop: vsView.TextAlign = taRightTop
    End Select
    If oFP.PropsFP.FontSize > 0 Then vsView.FontSize = oFP.PropsFP.FontSize
    If oFP.PropsFP.FontName <> "" Then
        vsView.FontName = oFP.PropsFP.FontName
    Else
        vsView.FontName = "Tahoma"
    End If
    vsView.FontBold = oFP.PropsFP.Bold
    'Para c/u de las posiciones las imprimo
    For iP = 1 To oFP.PropsFP.Posicion.Count
        If oFP.PropsFP.Width > 0 Then
            vValue = IIf(IsNull(vValue), "", vValue)
            vsView.TextBox vValue, oFP.PropsFP.Posicion.GetPosicion(iP).posX, oFP.PropsFP.Posicion.GetPosicion(iP).posY, oFP.PropsFP.Width, vsView.TextHeight(CStr(vValue)) + 40
        Else
            vsView.CurrentX = oFP.PropsFP.Posicion.GetPosicion(iP).posX
            vsView.CurrentY = oFP.PropsFP.Posicion.GetPosicion(iP).posY
            vsView.Paragraph = IIf(IsNull(vValue), "", vValue)
        End If
    Next
End Sub

Private Sub loc_PrintEtiquetas(ByVal vsView As vsPrinter)
    Dim iQ As Integer
    Dim oFP As clsRegFieldPrint
    For iQ = 1 To oFieldPrint.Count
        Set oFP = oFieldPrint.GetField(iQ)
        If oFP.Tipo = Label Then loc_PrintFieldLabel oFP, vsView, oFP.Nombre
    Next
End Sub

Private Sub loc_PrintField(ByVal vsView As vsPrinter, ByVal sNameField As String, ByVal vValue As Variant)
    Dim iQ As Integer
    Dim oFP As clsRegFieldPrint
    For iQ = 1 To oFieldPrint.Count
        Set oFP = oFieldPrint.GetField(iQ)
        If oFP.Tipo = Campo And LCase(oFP.Nombre) = LCase(sNameField) Then loc_PrintFieldLabel oFP, vsView, vValue
    Next
End Sub

Public Function PrintDocumento(ByVal sQy As String, ByRef vsView As vsPrinter, ByVal EnPapelBlanco As Boolean) As Boolean
On Error GoTo errPD
'Ya tuve que cargar el objeto con la info de impresión.
Dim rsP As rdoResultset
Dim iQ As Integer, iCol As Integer, sResult As String
Dim oFP As clsRegFieldPrint


    ReDim arrTPos(0)
    'Inicio el objeto
    If Not fnc_InitPrinter(vsView) Then Exit Function
               
    Set rsP = cBase.OpenResultset(sQy, rdOpenForwardOnly, rdConcurValues)
    If Not rsP.EOF Then
    
        Dim m_Tablas As New Collection
        
        'Primero imprimo los labels
        loc_PrintEtiquetas vsView
        
        If EnPapelBlanco Then
            vsView.FontBold = True
            vsView.CurrentX = 700
            vsView.CurrentY = 700
            vsView.Paragraph = "Copia en papel blanco"
            vsView.FontBold = False
        End If
        
        Dim sArticulosEnRemito As String
        ' Campoclave = ArtsEnRemito
        For iCol = 0 To rsP.rdoColumns.Count - 1
            If Trim(LCase(rsP(iCol).Name)) = "artsenremito" Then
                sArticulosEnRemito = Trim(rsP(iCol).Value)
                Exit For
            End If
        Next
        
        For iCol = 0 To rsP.rdoColumns.Count - 1
            'Ahora para todas las columnas del primer registro imprimo todos los valores campos.
            loc_PrintField vsView, rsP(iCol).Name, rsP(iCol).Value
        Next
        
        Dim oTabla As clsTablaRenglon
        'Me posiciono en el primer valor de una tabla.
        For iQ = 1 To oFieldPrint.Count
            Set oFP = oFieldPrint.GetField(iQ)
            If oFP.Tipo = Tabla Then
                For iCol = 1 To oFP.PropsFP.Posicion.Count
                    Set oTabla = New clsTablaRenglon
                    With oTabla
                        .IndexOrig = iQ
                        Set .PosInit = oFP.PropsFP.Posicion.GetPosicion(iCol)
                    End With
                    m_Tablas.Add oTabla
                    Set oTabla = Nothing
                Next
            End If
        Next
        
        Do While Not rsP.EOF
            'Para todas las tablas que cargue voy agregando sus renglones.
            For iQ = 1 To m_Tablas.Count
                Set oTabla = m_Tablas(iQ)
                sResult = oFieldPrint.GetField(oTabla.IndexOrig).PropsFP.TablaCampos
                'Recorro las columnas y cambio los valores
                For iCol = 0 To rsP.rdoColumns.Count - 1
                    sResult = Replace(sResult, "[" & LCase(Trim(rsP(iCol).Name)) & "]", IIf(IsNull(rsP(iCol).Value), "", rsP(iCol).Value), , , vbTextCompare)
                Next
                m_Tablas(iQ).AddRenglon sResult
            Next
            rsP.MoveNext
        Loop
        rsP.Close
        
        vsView.FontName = "Tahoma"
        
        'Ahora imprimo c/u de las tablas.
        For iQ = 1 To m_Tablas.Count
            Set oTabla = m_Tablas(iQ)
            vsView.CurrentX = oTabla.PosInit.posX
            vsView.CurrentY = oTabla.PosInit.posY
            vsView.TableBorder = tbNone
            If iQ = 2 And sArticulosEnRemito <> "" Then
                vsView.FontBold = True
                vsView.Paragraph = sArticulosEnRemito
            Else
                vsView.FontBold = False
                For iCol = 1 To oTabla.Count
                    vsView.AddTable oFieldPrint.GetField(oTabla.IndexOrig).PropsFP.Width, "", oTabla.GetRenglon(iCol)
                Next
            End If
        Next
        Set m_Tablas = Nothing
    End If
    vsView.EndDoc
'    vsView.Visible = True
'    vsView.Move 0, vsView.Top - 100, 6000, 6500
    vsView.PrintDoc False
    Debug.Print "Documento " & sQy
    PrintDocumento = True
    Exit Function
errPD:
    objGral.OcurrioError "Error al imprimir el documento.", Err.Description, "Print Manager"
End Function
