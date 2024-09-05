Attribute VB_Name = "ModSuperPosicion"
Option Explicit

Public Function superp_ValSuperposicion(Ind As Integer) As Double
'Asigna de Dom a Sáb Nos. de forma q se puedan superponer los días de la semana.
Dim ValAux As Double, I As Integer
ValAux = 1
For I = 2 To Ind
    ValAux = ValAux * 2
Next
superp_ValSuperposicion = ValAux
End Function

Public Function superp_MatrizSuperposicion(ByVal Valor As Long) As String
Dim VAux As Integer, TxtAux As String
superp_MatrizSuperposicion = ""
If Valor = 0 Then Exit Function
Do Until Valor = 0
    VAux = superp_IndSuperposicion(Valor)
    TxtAux = VAux & "," & TxtAux
    Valor = Valor - superp_ValSuperposicion(VAux)
Loop
superp_MatrizSuperposicion = Left(TxtAux, Len(TxtAux) - 1)
End Function

Public Function superp_IndSuperposicion(Valor As Long) As Integer
Dim ValAux As Double, Cant As Integer
ValAux = Valor
Do While True
    ValAux = ValAux / 2
    If ValAux >= 1 Then
        Cant = Cant + 1
    Else
        Exit Do
    End If
Loop
superp_IndSuperposicion = Cant + 1
End Function


