Attribute VB_Name = "modLocal"
Option Explicit

Type enuItems
    Moneda As Integer
    Tasa As Currency
    Rubro As String
    Neto As Currency
    Cofis As Currency
    IVA As Currency
    IDGasto As Long
End Type

Public arrTasas() As enuItems

Dim aIdx As Integer

Public Function arr_Agregar(idMoneda As Variant, idTasa As Currency, idRubro As String, _
                                    idNeto As Currency, idCofis As Currency, idIVA As Currency, IDGasto As Long) As Boolean
                                    
On Error GoTo errAdd
    arr_Agregar = False
    
    If arrTasas(0).Moneda <> 0 Then
        aIdx = UBound(arrTasas) + 1
        ReDim Preserve arrTasas(aIdx)
    Else
        aIdx = 0
    End If
    
    With arrTasas(aIdx)
        .Moneda = idMoneda
        .Tasa = idTasa
        .Rubro = Trim(idRubro)
        .Neto = idNeto
        .Cofis = idCofis
        .IVA = idIVA
        .IDGasto = IDGasto
    End With
    
    arr_Agregar = True
    Exit Function
    
errAdd:
End Function

Public Function arr_Sort() As Boolean
    'Ordena por Moneda y Tasa de Cambio
    On Error GoTo errSort
    arr_Sort = False
    
    If arrTasas(0).Moneda = 0 Then arr_Sort = True: Exit Function
    
    Dim arrAux() As enuItems
    ReDim Preserve arrAux(UBound(arrTasas))
    
    Dim mMoneda As Integer, mTC As Currency
    Dim mElem As Integer, mIdxAux As Integer
    
    mMoneda = arrTasas(0).Moneda
    mIdxAux = -1
    mTC = 999
    Do While arrTasas(0).Moneda <> 0
        mElem = -1
        mTC = 999
        
        For aIdx = 0 To UBound(arrTasas)
            
            If arrTasas(aIdx).Moneda = mMoneda And arrTasas(aIdx).Tasa <= mTC Then
                mTC = arrTasas(aIdx).Tasa
                mElem = aIdx
            End If

        Next
    
        If mElem = -1 And arrTasas(0).Moneda <> 0 Then  'Cambio de moneda
            mMoneda = arrTasas(0).Moneda
            
        Else
            If mElem <> -1 Then
                mIdxAux = mIdxAux + 1
                arrAux(mIdxAux) = arrTasas(mElem)
                
                BorroElemento mElem
            End If
        End If
    Loop
    
    arrTasas = arrAux
    ReDim arrAux(0)
    
    arr_Sort = True
    Exit Function
errSort:

End Function

Private Function BorroElemento(mIdBorrar As Integer)

Dim arrRet() As enuItems
Dim mI As Integer, mIE As Integer
    
    ReDim arrRet(UBound(arrTasas))
    For mI = 0 To UBound(arrTasas)
        
        If mI <> mIdBorrar Then
            arrRet(mIE) = arrTasas(mI)
            mIE = mIE + 1
        End If
        
    Next
    
    arrTasas = arrRet
    ReDim arrRet(0)
    
End Function

