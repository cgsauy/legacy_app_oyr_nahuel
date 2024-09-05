Attribute VB_Name = "modAuxiliar"
Option Explicit

Private Type typDisp        'Definicion para Disponiblidades
    Id As Long
    Nombre As String
    Moneda As Integer
    Bancaria As Boolean
End Type

Public arrDisp() As typDisp

Public prmFCierreIVA As Date

Public Function dis_StartArray()

On Error GoTo errCargo
Dim mQ As Integer: mQ = 0

    cons = "Select * from Disponibilidad Order by DisNombre"
    Set RsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        ReDim Preserve arrDisp(mQ)
        
        With arrDisp(mQ)
            .Id = RsAux!DisID
            .Nombre = Trim(RsAux!DisNombre)
            .Moneda = RsAux!DisMoneda
            .Bancaria = Not IsNull(RsAux!DisSucursal)
        End With
        
        mQ = mQ + 1
        RsAux.MoveNext
    Loop
    RsAux.Close

Exit Function
errCargo:
    clsGeneral.OcurrioError "Error al cargar array de disponibilidades.", Err.Description
End Function

Public Function dis_IdxArray(idDisponibilidad As Long) As Integer
    
    dis_IdxArray = -1
    
    For I = LBound(arrDisp) To UBound(arrDisp)
        If arrDisp(I).Id = idDisponibilidad Then
            dis_IdxArray = I
            Exit For
        End If
    Next
    
End Function

Public Function dis_CargoDisponibilidades(mControl As Control, mMoneda As Integer)

On Error GoTo errCargo
Dim mOldSel As Integer

    mOldSel = mControl.ListIndex
    mControl.Clear
    
    For I = LBound(arrDisp) To UBound(arrDisp)
        If arrDisp(I).Moneda = mMoneda Then
            mControl.AddItem Trim(arrDisp(I).Nombre)
            mControl.ItemData(mControl.NewIndex) = arrDisp(I).Id
        End If
    Next
    
    mControl.AddItem "(Otras)", 0
    mControl.ItemData(0) = 0
    
    If mMoneda = paMonedaPesos Then
        BuscoCodigoEnCombo mControl, paDisponibilidad
    Else
        mControl.ListIndex = mOldSel
    End If
Exit Function
errCargo:
    clsGeneral.OcurrioError "Error al cargar array de disponibilidades.", Err.Description
End Function

Public Function dis_FechaCierre(idDisponibilidad As Long, dAnteriorA As Date) As Date
On Error GoTo errQuery
Dim rsLoc As rdoResultset

    dis_FechaCierre = CDate("1/1/1900")
    cons = "Select Top 2 * " & _
                " From SaldoDisponibilidad " & _
                " Where SDiFecha >= '" & Format(dAnteriorA, "mm/dd/yyyy") & "'" & _
                " And SDiDisponibilidad = " & idDisponibilidad & _
                " Order by SDiFecha ASC"
   
    Set rsLoc = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsLoc.EOF Then
        Dim mHora As String
        dis_FechaCierre = rsLoc!SDiFecha
        mHora = rsLoc!SDiHora
        
        'Por los saldos iniciales de disponibilidades
        If Format(rsLoc!SDiFecha, "dd/mm/yyyy") = Format(dAnteriorA, "dd/mm/yyyy") And rsLoc!SDiHora = "00:00:00" Then
            rsLoc.MoveNext
            If Not rsLoc.EOF Then
                dis_FechaCierre = rsLoc!SDiFecha
                mHora = rsLoc!SDiHora
            End If
        End If
        
        If mHora = "00:00:00" Then dis_FechaCierre = dis_FechaCierre - 1
        
    End If
    rsLoc.Close

errQuery:
End Function

Public Sub dis_BorroRelacionCheque(mIDCheque As Long, mIDCompra As Long)

Dim rsChk As rdoResultset
    
    cons = "Delete ChequePago Where CPaIDCheque = " & mIDCheque & " And CPaIDCompra = " & mIDCompra
    cBase.Execute cons
            
    'Si no hay mas relaciones de pago borro el cheque
    cons = "Select * from ChequePago Where CPaIDCheque = " & mIDCheque
    Set rsChk = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If rsChk.EOF Then
        cons = "Delete Cheque Where CheId = " & mIDCheque
        cBase.Execute cons
    End If
    rsChk.Close
 
End Sub
