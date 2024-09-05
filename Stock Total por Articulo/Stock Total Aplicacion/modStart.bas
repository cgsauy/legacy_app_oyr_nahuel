Attribute VB_Name = "modStart"
Option Explicit

Public Sub main()
Dim sParam As String
    
    sParam = Trim(Command())
    Dim objStock As New clsStockTotal
    If IsNumeric(sParam) Then
        objStock.ShowStockTotal Val(sParam)
    Else
        objStock.ShowStockTotal
    End If
    Set objStock = Nothing
    End

End Sub
