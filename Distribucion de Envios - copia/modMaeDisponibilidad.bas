Attribute VB_Name = "modMaeDisponibilidad"
' ****** Módulo Maestro de Disponibilidades ****** --------------------------------------------------------
' Rutinas para manejo de disponibilidades en moneda extranjera y moneda por defecto de sucursales

Option Explicit
Private rsMae As rdoResultset
Private mSQL As String

'DATOS DE LAS MONEDAS -------------------------------------------------------------------
Private Type typMoneda
    Codigo As Integer
    CoeficienteMora As Currency
    Redondeo As String
End Type

Public Enum enuMoneda
    pCodigo = 1
    pCoeficienteMora = 2
    pRedondeo = 3
End Enum

Private arrMonedas() As typMoneda
'----------------------------------------------------------------------------------------------------

'   Retorna las sucursales que manejen la disponibilidad xxx.
'   Chequea los campos SucDisponibilidad y  SucDisponibilidadME
'   -> Retorna string separado por comas con  ids de sucursales
Public Function dis_SucursalesConDisponibilidad(mIDDisponibilidad As Long) As String
On Error GoTo errFnc
Dim I As Integer

    dis_SucursalesConDisponibilidad = ""
    
    Dim mRet() As Variant, mData() As String
    ReDim Preserve mRet(0): mRet(0) = 0
    Dim bOK As Boolean
    
    mSQL = "Select SucCodigo, SucDisponibilidad, SucDisponibilidadME From Sucursal"
    Set rsMae = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    Do While Not rsMae.EOF
        bOK = False
        If Not IsNull(rsMae!SucDisponibilidad) Then
            If rsMae!SucDisponibilidad = mIDDisponibilidad Then
                If mRet(0) <> 0 Then ReDim Preserve mRet(UBound(mRet) + 1)
                mRet(UBound(mRet)) = rsMae!SucCodigo
                bOK = True
            End If
        End If
        
        If Not bOK Then
            If Not IsNull(rsMae!SucDisponibilidadME) Then
                If Trim(rsMae!SucDisponibilidadME) <> "" Then
                    mData = Split(Trim(rsMae!SucDisponibilidadME), ",")
                    For I = LBound(mData) To UBound(mData)
                        If CLng(mData(I)) = mIDDisponibilidad Then
                            If mRet(0) <> 0 Then ReDim Preserve mRet(UBound(mRet) + 1)
                            mRet(UBound(mRet)) = rsMae!SucCodigo
                            Exit For
                        End If
                    Next
                    
                End If
            End If
        End If
        rsMae.MoveNext
    Loop
    rsMae.Close
    
    If mRet(0) <> 0 Then dis_SucursalesConDisponibilidad = Join(mRet, ",")
        
errFnc:
End Function

'   Retorna la Disponibilidad que corresponde a la Sucursal y Moneda
'   en la que se va a hacer el movimiento
Public Function dis_DisponibilidadPara(mIDSucursal As Long, mIDMoneda As Long) As Long

On Error GoTo errFnc

    dis_DisponibilidadPara = 0
    
    Dim mRet As Long
    Dim mIds As String
    mIds = ""
    
    mSQL = "Select SucDisponibilidad, SucDisponibilidadME " & _
                " From Sucursal Where SucCodigo = " & mIDSucursal
    Set rsMae = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    If Not rsMae.EOF Then
        
        If Not IsNull(rsMae!SucDisponibilidad) Then mIds = rsMae!SucDisponibilidad
        
        If Not IsNull(rsMae!SucDisponibilidadME) Then
            If Trim(rsMae!SucDisponibilidadME) <> "" Then
                If mIds <> "" Then mIds = mIds & ","
                mIds = mIds & Trim(rsMae!SucDisponibilidadME)
            End If
        End If
        
    End If
    rsMae.Close
    
    If mIds <> "" Then
        mSQL = "Select * From Disponibilidad Where DisID In (" & mIds & ")"
        Set rsMae = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
        Do While Not rsMae.EOF
            If rsMae!DisMoneda = mIDMoneda Then
                dis_DisponibilidadPara = rsMae!DisID
                Exit Do
            End If
            rsMae.MoveNext
        Loop
        rsMae.Close
    End If
    
errFnc:
End Function

'-----------------------------------------------------------------------------------------------------------------------
'   Carga array con los datos de las monedas
Public Function dis_CargoArrayMonedas() As Boolean

On Error GoTo errMonedas
    ReDim Preserve arrMonedas(0)
    
    Cons = "Select * from Moneda"
    Set rsMae = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsMae.EOF
        If arrMonedas(0).Codigo <> 0 Then ReDim Preserve arrMonedas(UBound(arrMonedas) + 1)
        With arrMonedas(UBound(arrMonedas))
            .Codigo = rsMae!MonCodigo
            
            .CoeficienteMora = 1
            If Not IsNull(rsMae!MonCoeficienteMora) Then
                .CoeficienteMora = ((rsMae!MonCoeficienteMora / 100) + 1) ^ (1 / 30)                         'Como es mensual calculo el diario
            End If
            
            .Redondeo = 1
            If Not IsNull(rsMae!MonRedondeo) Then .Redondeo = Trim(rsMae!MonRedondeo)
        End With
        
        rsMae.MoveNext
    Loop
    rsMae.Close
    
Exit Function
errMonedas:
    MsgBox "Error al cargar el array de monedas. Los cálculos de intereses y redondeos pueden dar ERROR." & vbCr & vbCr & "Error:" & Err.Description, vbCritical, "Error Con Parámetros !!!"
End Function

'-----------------------------------------------------------------------------------------------------------------------
'   Retorna propiedad para el id de moneda. Debe estar cargado el array !!!!
Public Function dis_arrMonedaProp(mIDMoneda As Long, idProp As enuMoneda) As Variant
    On Error GoTo errArray
    Dim idx As Integer
    
    dis_arrMonedaProp = -1
    
    For idx = LBound(arrMonedas) To UBound(arrMonedas)
        If arrMonedas(idx).Codigo = mIDMoneda Then
            Select Case idProp
                Case enuMoneda.pCodigo: dis_arrMonedaProp = arrMonedas(idx).Codigo
                Case enuMoneda.pCoeficienteMora: dis_arrMonedaProp = arrMonedas(idx).CoeficienteMora
                Case enuMoneda.pRedondeo: dis_arrMonedaProp = arrMonedas(idx).Redondeo
            End Select
            Exit For
        End If
    Next
errArray:
End Function

Public Function Redondeo(Valor As Currency, mPatron As String) As String
On Error GoTo errRND
Dim Numero As Currency

    Numero = Abs(Valor)
    
    'Redondeo = Format(Numero, "#,##0") & ".00"
    mPatron = Trim(mPatron)
    
    Dim mRet As String
    
    Dim mDigito As Integer      '1 ó 5
    Dim aQDec As Integer: aQDec = 0
    
    Dim aQDecNro As Integer: aQDecNro = 0
    If InStr(Numero, ".") <> 0 Then aQDecNro = Len(Mid(Numero, InStr(Numero, ".") + 1)) + 1 'Q decimales + el pto
    
    If InStr(mPatron, ".") <> 0 Then    'Redondeo c/Decimales ---------------------------------------------
        mDigito = Right(mPatron, 1)
        
        aQDec = Len(Mid(mPatron, InStr(mPatron, ".") + 1))     'Q decimales
        mRet = Format(Numero, "0." & String(aQDec, "0"))       'Formateo a Q decimales
        
        If mDigito = 5 Then                             'Corrigo a 0, 5 ó 10
            Select Case Right(mRet, 1)
                Case Is <= 2:           'Corrigo A 0
                    mRet = Mid(mRet, 1, Len(mRet) - 1) & "0"
                
                Case Is <= 7            'Corrigo A 5
                    mRet = Mid(mRet, 1, Len(mRet) - 1) & "5"
                    
                Case Is <= 9            'Corrigo a 10
                    Dim aDiff As Integer
                    aDiff = 10 - CInt(Right(mRet, 1))
                    mRet = mRet + CCur("0." & String(aQDec - 1, "0") & aDiff)
            End Select
        End If
    
    Else                'Redondeo sin decimales     ----------------------------------------------------------------
        mDigito = Left(mPatron, 1)
        mRet = Numero
        
        Dim mARd As Currency, mCompare As Long
        Dim longFmt As Long
        longFmt = CLng(mPatron)
        
        If mDigito = 5 Then
            mARd = Right(mRet, Len(mPatron) + aQDecNro)
            mCompare = CLng("1" & String(Len(mPatron), "0"))
            
            Select Case (mCompare - mARd)
                Case Is <= (longFmt / 2)
                    mRet = mRet + mCompare - mARd
                    
                Case Is <= longFmt + (longFmt / 2)
                    If longFmt < mARd Then
                        mRet = CCur(mRet) - (mARd - longFmt)
                    Else
                        mRet = CCur(mRet) + (longFmt - mARd)
                    End If
                
                Case Else
                    mRet = CCur(mRet) - mARd
            End Select
        
        Else
            
            mARd = Val(Right(mRet, Len(mPatron) - 1 + aQDecNro))
            mCompare = CLng(mPatron)

            Select Case (mCompare - mARd)
                Case Is <= (longFmt / 2)
                    mRet = CCur(mRet) + mCompare - mARd
                    
                Case Else
                    mRet = CCur(mRet) - mARd
            End Select
        End If
        
    End If
    
    If Valor < 0 Then mRet = CCur(mRet) * -1
    If aQDec <> 0 Then
        Redondeo = Format(mRet, "#,##0." & String(aQDec, "0"))
    Else
        Redondeo = Format(mRet, "#,##0")
    End If
    Exit Function
    
errRND:
    If Valor < 0 Then Numero = Numero * -1
    Redondeo = Numero
End Function

