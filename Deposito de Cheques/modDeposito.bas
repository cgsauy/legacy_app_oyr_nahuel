Attribute VB_Name = "modDeposito"
Option Explicit

Public Enum enuPrmsDepositos
    TipoTranferencia = 1
    SRGasto = 2
    DispDeposito = 3
    DispDepositoFletes = 4
End Enum

'Definicion de Tipos para registro de Gastos    ----------------
Private Type typGasto
    IdTipoTransferencia As Long
    ImporteAlDia As Currency
    IdDisponibilidadSalida As Long
    IdDisponibilidadEntrada As Long
    
    idMoneda As Long
    
    ImporteDiferido As Currency
    IdRubroSalida As Long
    IdSubrubroSalida As Long
    NameRSalida As String
    NameSRSalida As String
    
    SucursalNombre As String
    SucursalID As Long
    IdProveedorGasto As Long
    zPrmsDepositos As String
End Type

Public dGastos() As typGasto
'--------------------------------------------------------------------

Public Function arrG_AddItem(idSucursal As Long, nameSucursal As String, Importe As Currency, AlDia As Boolean, mMoneda As Long, ByVal Tag As Byte) As Boolean
On Error GoTo errAdd
    arrG_AddItem = True
    
    Dim idx As Integer, bAddOk As Boolean
    bAddOk = False
    
    For idx = LBound(dGastos) To UBound(dGastos)
                
        If dGastos(idx).SucursalID = idSucursal Then
            If AlDia Then
                dGastos(idx).ImporteAlDia = dGastos(idx).ImporteAlDia + Importe
            Else
                dGastos(idx).ImporteDiferido = dGastos(idx).ImporteDiferido + Importe
            End If
            bAddOk = True: Exit For
        End If
        
    Next
    
    If Not bAddOk Then
        If UBound(dGastos) = 0 And dGastos(0).SucursalID = 0 Then
            idx = 0
        Else
            idx = UBound(dGastos) + 1
            ReDim Preserve dGastos(idx)
        End If
        
        With dGastos(idx)
            .SucursalID = idSucursal
            .SucursalNombre = Trim(nameSucursal)
            .IdTipoTransferencia = 0
            .ImporteAlDia = 0
            .IdDisponibilidadSalida = paDisponibilidad
            .idMoneda = mMoneda
            .ImporteDiferido = 0
            If AlDia Then .ImporteAlDia = Importe Else .ImporteDiferido = Importe
                        
            .zPrmsDepositos = ""
            
            Cons = "Select * from SucursalDeBanco Where SBaCodigo = " & .SucursalID
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                .IdProveedorGasto = RsAux!SBaBanco
                If Not IsNull(RsAux!SBaPrmsDepositos) Then .zPrmsDepositos = Trim(RsAux!SBaPrmsDepositos)
            End If
            RsAux.Close
            
            .IdTipoTransferencia = get_PrmsDepositos(.zPrmsDepositos, TipoTranferencia)
            If Tag = 0 Then
                'Comercio
                .IdDisponibilidadEntrada = get_PrmsDepositos(.zPrmsDepositos, DispDeposito)
            Else
                'Fleteros
                .IdDisponibilidadEntrada = get_PrmsDepositos(.zPrmsDepositos, DispDepositoFletes)
            End If
            .IdSubrubroSalida = get_PrmsDepositos(.zPrmsDepositos, SRGasto)
            
            If .IdSubrubroSalida <> 0 Then
                Cons = "Select * from SubRubro, Rubro " & _
                            " Where SRuID = " & .IdSubrubroSalida & _
                            " And SRuRubro = RubID "
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsAux.EOF Then
                    .IdRubroSalida = RsAux!RubID
                    .NameRSalida = Trim(RsAux!RubNombre)
                    .NameSRSalida = Trim(RsAux!SRuNombre)
                End If
                RsAux.Close
            End If
            
        End With
    End If
    Exit Function
    
errAdd:
    arrG_AddItem = False
End Function

Public Function put_PrmsDepositos(mPrms As String, mIDPrm As enuPrmsDepositos, mValor As Long) As String
On Error GoTo errPut

    put_PrmsDepositos = mPrms
    Dim mValorPrm As String
    mValorPrm = get_PrmsName(mIDPrm)
    
    If Trim(mPrms) = "" Then
        put_PrmsDepositos = mValorPrm & ":" & mValor
        Exit Function
    End If
    
    Dim mArray() As String, mValue() As String, idx As Integer
    Dim bOK As Boolean: bOK = False
    
    mArray = Split(mPrms, "|")
        
    For idx = LBound(mArray) To UBound(mArray)
        mValue = Split(mArray(idx), ":")
        If LCase(Trim(mValue(0))) = LCase(mValorPrm) Then
            mValue(1) = mValor
            mArray(idx) = Join(mValue, ":")
            bOK = True
            Exit For
        End If
    Next
    
    If Not bOK Then
        ReDim Preserve mArray(UBound(mArray) + 1)
        idx = UBound(mArray)
        mArray(idx) = mValorPrm & ":" & mValor
    End If
    put_PrmsDepositos = VBA.Join(mArray, "|")
    
errPut:
End Function

Private Function get_PrmsName(mIDPrm As enuPrmsDepositos) As String

    Select Case mIDPrm
        Case enuPrmsDepositos.DispDeposito: get_PrmsName = "DDE"
        Case enuPrmsDepositos.SRGasto: get_PrmsName = "SRG"
        Case enuPrmsDepositos.TipoTranferencia: get_PrmsName = "TTR"
        Case enuPrmsDepositos.DispDepositoFletes: get_PrmsName = "Tag1"
    End Select
    
End Function

Public Function get_PrmsDepositos(mPrms As String, mValor As enuPrmsDepositos) As Variant
On Error GoTo errGetValor

    get_PrmsDepositos = 0
    If Trim(mPrms) = "" Then Exit Function
    
    Dim mStr As String: mStr = Trim(mPrms)
    Dim mArray() As String, mValue() As String, idx As Integer
    
    mArray = Split(mStr, "|")
    mStr = get_PrmsName(mValor)
    
    For idx = LBound(mArray) To UBound(mArray)
        mValue = Split(mArray(idx), ":")
        If LCase(Trim(mValue(0))) = LCase(mStr) Then
            get_PrmsDepositos = mValue(1)
            Exit For
        End If
    Next

errGetValor:
End Function


Public Function ing_BuscoSubrubro(mControlR As TextBox, mControlSR As TextBox) As Boolean
On Error GoTo errBS

    ing_BuscoSubrubro = False
    Dim aQ As Integer, aID As Long, aTexto As String
    aQ = 0: aID = 0
    
    mControlSR.Text = Replace(RTrim(mControlSR.Text), " ", "%")
    
    Cons = "Select SRuID, SRuNombre as 'SubRubro', SRuCodigo as 'Cód. SR', RubNombre as 'Rubro', RubCodigo as 'Cód. Rubro'" _
            & " from SubRubro, Rubro " _
            & " Where SRuNombre like '" & Trim(mControlSR.Text) & "%'" _
            & " And SRuRubro = RubID "
            '& " And SRuCodigo Not like '" & paRubroDisponibilidad & "%'"
    If Val(mControlR.Tag) <> 0 Then Cons = Cons & " And RubID= " & Val(mControlR.Tag)
    Cons = Cons & " Order by SRuNombre"
                
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        aQ = 1: aID = RsAux!SRuID: aTexto = Trim(RsAux(1))
        RsAux.MoveNext
        If Not RsAux.EOF Then
            aQ = 2: aID = 0
        End If
    End If
    RsAux.Close
    
    Select Case aQ
        Case 0: MsgBox "No existen Subrubros para el texto ingresado.", vbExclamation, "No hay datos"
        
        Case 1:
                mControlSR.Text = aTexto: mControlSR.Tag = aID
        
        Case 2:
                Dim aLista As New clsListadeAyuda
                aID = aLista.ActivarAyuda(cBase, Cons, 5500, 1, "Sub Rubros")
                
                If aID <> 0 Then
                    aTexto = Trim(aLista.RetornoDatoSeleccionado(1))
                    aID = aLista.RetornoDatoSeleccionado(0)
                End If
                
                Set aLista = Nothing
    End Select
    
    If aID <> 0 Then
        Cons = "Select * from Subrubro, Rubro " & _
                   " Where SRuID = " & aID & _
                   " And SRuRubro = RubID"
        
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            mControlR.Text = Trim(RsAux!RubNombre)
            mControlR.Tag = RsAux!SRuRubro
            
            mControlSR.Text = Trim(RsAux!SRuNombre)
            mControlSR.Tag = RsAux!SRuID
            ing_BuscoSubrubro = True
        End If
        RsAux.Close
        
    End If
    
    Screen.MousePointer = 0
    Exit Function

errBS:
    clsGeneral.OcurrioError "Error al buscar el subrubro.", Err.Description
    Screen.MousePointer = 0
End Function


Public Function ing_BuscoRubro(mControlR As TextBox) As Boolean
On Error GoTo errBS

    ing_BuscoRubro = False
    Dim aQ As Integer, aID As Long, aTexto As String
    aQ = 0: aID = 0
    
    mControlR.Text = Replace(RTrim(mControlR.Text), " ", "%")
    
    Cons = "Select RubID, RubNombre as 'Rubro', RubCodigo as 'Cód. Rubro'" _
            & " from Rubro " _
            & " Where RubNombre like '" & Trim(mControlR.Text) & "%'" _
            & " Order by RubNombre"
                
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        aQ = 1: aID = RsAux!RubID: aTexto = Trim(RsAux(1))
        RsAux.MoveNext
        If Not RsAux.EOF Then
            aQ = 2: aID = 0
        End If
    End If
    RsAux.Close
    
    Select Case aQ
        Case 0: MsgBox "No existen rubros para el texto ingresado.", vbExclamation, "No hay datos"
        
        Case 1:
                mControlR.Text = aTexto: mControlR.Tag = aID
        
        Case 2:
                Dim aLista As New clsListadeAyuda
                aID = aLista.ActivarAyuda(cBase, Cons, 4500, 1, "Rubros")
                If aID <> 0 Then
                    aTexto = Trim(aLista.RetornoDatoSeleccionado(1))
                    aID = aLista.RetornoDatoSeleccionado(0)
                End If
                Set aLista = Nothing
    End Select
    
    If aID <> 0 Then
        mControlR.Text = aTexto
        mControlR.Tag = aID
            
        Cons = "Select Top 2 * from Subrubro" & _
                   " Where SRuRubro = " & aID '& _
                   " And SRuCodigo Not like '" & paRubroDisponibilidad & "%'"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            aTexto = Trim(RsAux!SRuNombre)
            aID = RsAux!SRuID
            RsAux.MoveNext
            If RsAux.EOF Then
                mControlR.Text = aTexto
                mControlR.Tag = aID
            End If
        End If
        RsAux.Close
        
    End If
    
    Screen.MousePointer = 0
    Exit Function

errBS:
    clsGeneral.OcurrioError "Error al buscar el rubro.", Err.Description
    Screen.MousePointer = 0
End Function


