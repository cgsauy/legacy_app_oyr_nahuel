Attribute VB_Name = "modExplorer"
Option Explicit

Type enuItems
    Codigo As Long
    Tipo As Integer
    Nombre As String
End Type

Public arrClientes() As enuItems

Dim aIdx As Integer
Private Const maxArray = 26

Public Function arr_AddItem(ByVal lID As Long, ByVal sName As String)
    
    '7/4/2008 si el artículo lo tengo en la lista --> lo envío a la última pos del array.
    'Busco en el array si lo tengo --> lo elimino.

'    If arrcli_Item(lID) >= 0 Then Exit Function
    
    For aIdx = LBound(arrClientes) To UBound(arrClientes)
        If arrClientes(aIdx).Codigo = lID Then
            
            With arrClientes(aIdx)
                .Codigo = 0
                .Nombre = ""
                .Tipo = 0
            End With
        
            Exit For
        End If
    Next
    
    
    aIdx = arrcli_MaxItem
    If aIdx = maxArray Then
        arrcli_CorrijoPos
        aIdx = maxArray - 1
    End If
    
    With arrClientes(aIdx)
        .Codigo = lID
        .Nombre = Trim(sName)
    End With
    
End Function

Public Function arrcli_Item(idCliente As Long) As Long
'Retorna la posicion del array
    On Error GoTo errBuscar
    
    arrcli_Item = -1
    For aIdx = LBound(arrClientes) To UBound(arrClientes)
        If arrClientes(aIdx).Codigo = idCliente Then
            arrcli_Item = aIdx
            Exit For
        End If
    Next
    
errBuscar:
End Function

Private Function arrcli_MaxItem() As Integer
'Retorna la posicion del maximo elemento cargado
    On Error GoTo errBuscar
    
    arrcli_MaxItem = maxArray
    For aIdx = LBound(arrClientes) To UBound(arrClientes)
        If arrClientes(aIdx).Codigo = 0 Then
            arrcli_MaxItem = aIdx
            Exit For
        End If
    Next
    
errBuscar:
End Function

Private Function arrcli_CorrijoPos() As Integer

    On Error GoTo errBuscar
    'El q se va es el primero

    For aIdx = LBound(arrClientes) + 1 To UBound(arrClientes)
        arrClientes(aIdx - 1).Codigo = arrClientes(aIdx).Codigo
        arrClientes(aIdx - 1).Nombre = arrClientes(aIdx).Nombre
        arrClientes(aIdx - 1).Tipo = arrClientes(aIdx).Tipo
    Next
    
    aIdx = UBound(arrClientes)
    arrClientes(aIdx).Codigo = 0
    arrClientes(aIdx).Nombre = ""
    arrClientes(aIdx).Tipo = 0
    
errBuscar:
End Function
