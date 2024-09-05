Attribute VB_Name = "modListados"
Option Explicit

'------------------------------------------------------------------------------------------------------------------------------------
'   Setea la impresora pasada como parámetro como: por defecto
'------------------------------------------------------------------------------------------------------------------------------------
Public Sub SeteoImpresoraPorDefecto(DeviceName As String)
Dim X As Printer

    For Each X In Printers
        Debug.Print X.DeviceName
        If Trim(X.DeviceName) = Trim(DeviceName) Then
            Set Printer = X
            Exit For
        End If
    Next
    
End Sub

Public Sub CargoDatosZoom(cCombo As Control, Optional Maximo As Integer = 120)

    Dim aValor As Integer
    aValor = 10
    Do While aValor <= 120
        cCombo.AddItem aValor
        aValor = aValor + 10
    Loop
    cCombo.ListIndex = 0
    
End Sub

Public Sub EnumeroPiedePagina(vsPrint As Control)
Dim I As Integer

    With vsPrint
        .Columns = 1
        .FontSize = 9: .Font = "Arial": .FontItalic = False: .FontBold = False
        .TextAlign = 2
        For I = 1 To .PageCount
            .StartOverlay I
            
            .CurrentX = .MarginLeft
            .CurrentY = .PageHeight - .MarginBottom + 150
            vsPrint = "Página " & Format(I) & " de " & Format(.PageCount)
            
            .EndOverlay
        Next
        .TextAlign = 0 'taLeft
    End With
    
End Sub

Public Function FormateoString(vsPrint As Control, cLargoCelda As Currency, strTexto As String) As String
Dim largo As Currency

    largo = vsPrint.TextWidth(strTexto)
    Do While vsPrint.TextWidth(strTexto) > cLargoCelda
        strTexto = Mid(strTexto, 1, Len(strTexto) - 1)
    Loop
    FormateoString = strTexto
    
End Function

Public Sub Zoom(vsPrint As Control, Valor As Integer)

     On Error Resume Next
    Screen.MousePointer = 11
    vsPrint.Visible = False
    vsPrint.Zoom = Val(Valor)
    Screen.MousePointer = 0
    vsPrint.Visible = True
    
End Sub

Public Sub ZoomIn(vsPrint As Control)
        
    On Error Resume Next
    Screen.MousePointer = 11
    vsPrint.Visible = False
    If vsPrint.Zoom > 10 Then vsPrint.Zoom = vsPrint.Zoom - 10
    vsPrint.Visible = True
    Screen.MousePointer = 0
    
End Sub

Public Sub ZoomOut(vsPrint As Control)
        
    On Error Resume Next
    Screen.MousePointer = 11
    vsPrint.Visible = False
    If vsPrint.Zoom < 120 Then vsPrint.Zoom = vsPrint.Zoom + 10
    vsPrint.Visible = True
    Screen.MousePointer = 0
    
End Sub

Public Sub IrAPagina(vsPrint As Control, Pagina As Integer)
    vsPrint.PreviewPage = Pagina
End Sub

Public Sub EncabezadoListado(vsPrint As Control, strTitulo As String, sNombreEmpresa As Boolean)
    
    With vsPrint
        .HdrFont = "Arial"
        .HdrFontSize = 10
        .HdrFontBold = False
    End With
    
    If sNombreEmpresa Then
        vsPrint.Header = strTitulo + "||Carlos Gutiérrez S.A."
    Else
        vsPrint.Header = strTitulo
    End If
    vsPrint.HdrFontBold = False: vsPrint.FontBold = False
    vsPrint.HdrFontSize = 10: vsPrint.Footer = Format(Now, "dd/mm/yy hh:mm")
    
End Sub

Public Function VerificoQueExistaImpresora(NombreImpresora As String) As Boolean
On Error GoTo ErrVQEI
Dim X As Printer
    VerificoQueExistaImpresora = False
    For Each X In Printers
        If Trim(X.DeviceName) = Trim(NombreImpresora) Then
            VerificoQueExistaImpresora = True
            Exit For
        End If
    Next
    Exit Function
ErrVQEI:
    MsgBox "Ocurrió un error al verificar que existan las impresoras asignadas. " & Err.Description
End Function
