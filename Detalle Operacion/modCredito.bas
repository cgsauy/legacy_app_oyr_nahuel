Attribute VB_Name = "modCredito"

'-------------------------------------------------------------------------------------------------------------------
'   Maneja el string del cumplimiento en el pago de cuotas.
'   Parametro Cumplimiento:
'       F: Solamente para dar formato (todos puntos)
'       P: Para dar formato a todas como pagas (todos ceros)
'-------------------------------------------------------------------------------------------------------------------
Public Function FormatoCumplimiento(Cuotas As Integer, Vencimiento As Date, Cumplimiento As String, Optional dFPago As Date = "01/01/1800")

Dim aInsertar As String
Dim aResta As Long
Dim aDifDias As Currency
Const sCumplimiento = "0123456789ABCDEFGHIJKLMNOPQRSTX"

    'El Valor Cuotas debe venir si tiene entrega con + 1 y Se usa para DAR FORMATO (Tantos . como cuotas)
    'Si en el parametro Cumplimieto viene una F es solamente para dar formato
    Select Case Trim(Cumplimiento)
        Case "F"
            FormatoCumplimiento = String(Cuotas, ".")
            Exit Function
        Case "P"
            FormatoCumplimiento = String(Cuotas, "0")
            Exit Function
    
        Case "": Cumplimiento = String(Cuotas, ".")
    End Select
    
    If dFPago = CDate("01/01/1800") Then    'Si no viene fecha de pago, calculo vencimiento con dia de hoy
        aDifDias = Vencimiento - gFechaServidor
    Else
        aDifDias = Vencimiento - dFPago
    End If
    
    'Select Case Vencimiento - gFechaServidor
    Select Case aDifDias
        Case Is < 0
            'aResta = Abs(Vencimiento - gFechaServidor)
            aResta = Abs(aDifDias)
            
            aResta = (aResta \ 10) + 1
            aInsertar = Trim(Mid(sCumplimiento, aResta, 1))
            If aInsertar = "" Then aInsertar = "X"
        
        Case Is >= 0
            aInsertar = "0"
    End Select
    
    FormatoCumplimiento = Left(Cumplimiento, InStr(Cumplimiento, ".") - 1) & Trim(aInsertar) & Trim(Mid(Cumplimiento, InStr(Cumplimiento, ".") + 1, Len(Cumplimiento)))
    
End Function

'-------------------------------------------------------------------------------------------------------------------
'   Coloca "V" a las cuotas vencidas en el string de cumplimiento (Reemplaza los puntos por V)
'-------------------------------------------------------------------------------------------------------------------
Public Function CumplimientoConVencimientos(Texto As String, FechaDoc As Date, DiasEn As Variant, _
                                                                   DiasCu As Variant, Distancia As Variant) As String

Dim aTexto As String
Dim Acumulado As Long
Dim iI As Integer

    aTexto = ""
    Acumulado = 0
    For iI = 1 To Len(Texto)
        If Mid(Texto, iI, 1) = "." Then
            Select Case iI
                Case 1          'Primera Cuota (Entrega o Cuota 1)-------------------------------------------------------------------
                    If Not IsNull(DiasEn) Then
                        If DateDiff("d", FechaDoc + DiasEn, Now) > 0 Then aTexto = aTexto & "V" Else: aTexto = aTexto & "."
                    Else
                        If Not IsNull(DiasCu) Then
                            If DateDiff("d", FechaDoc + DiasCu, Now) > 0 Then aTexto = aTexto & "V" Else: aTexto = aTexto & "."
                        End If
                    End If
                
                Case 2          '2a Cuota (Si hay entrega es la primera cuota, sino la 2a normal)--------------------------------
                    If Not IsNull(DiasCu) And Not IsNull(DiasEn) Then
                        If DateDiff("d", FechaDoc, Now) > DiasCu Then aTexto = aTexto & "V" Else: aTexto = aTexto & "."
                        Acumulado = DiasCu
                    Else
                        If Not IsNull(Distancia) Then
                            Acumulado = Acumulado + Distancia
                            If DateDiff("d", FechaDoc + Acumulado, Now) > 0 Then aTexto = aTexto & "V" Else: aTexto = aTexto & "."
                            'If DateDiff("d", FechaDoc + Acumulado, Now) > 0 Then aTexto = aTexto & "V" Else: aTexto = aTexto & "."
                        End If
                    End If
                
                Case Else       'El resto de las cuotas--------------------------------------------------------------------------------
                        If Not IsNull(Distancia) Then
                            Acumulado = Acumulado + Distancia
                            If DateDiff("d", FechaDoc + Acumulado, Now) > 0 Then aTexto = aTexto & "V" Else: aTexto = aTexto & "."
                            'If DateDiff("d", FechaDoc + Acumulado, Now) > 0 Then aTexto = aTexto & "V" Else: aTexto = aTexto & "."
                        End If
            End Select
        Else
            aTexto = aTexto & Mid(Texto, iI, 1)
            Select Case iI
                Case 2:
                    If Not IsNull(DiasCu) And Not IsNull(DiasEn) Then
                        Acumulado = DiasCu
                    Else
                        If Not IsNull(Distancia) Then Acumulado = Acumulado + Distancia
                    End If
                Case Is > 2: Acumulado = Acumulado + Distancia
            End Select
        End If
    Next iI
    
    CumplimientoConVencimientos = aTexto
    
End Function

'Cambie el coeficiente x un doble el 30/01/2003
Public Function CalculoMora(SaldoCuota As Currency, FechaDesde As Date, MoraACuenta As Currency, mCoeficiente As Double) As Currency

Dim Dias As Integer         'Cantidad de Dias de Mora
Dim mValor As Currency

    Dias = DateDiff("d", FechaDesde, gFechaServidor)
    Set rsQ = cBase.OpenResultset("SELECT dbo.ValorCuota(" & CCur(SaldoCuota) & ", " & Dias & ", Null, Null)", rdOpenDynamic, rdConcurValues)
    If Not rsQ.EOF Then
        'Le resto el valor de la cuota ya que la función lo suma
        CalculoMora = rsQ(0) - SaldoCuota
    End If
    rsQ.Close
    
'    mValor = (mCoeficiente ^ Dias) - 1     'Elevo a la cantidad de dias y le resto un peso
'    'Multiplico  por el valor de la cuota
'    CalculoMora = (mValor * SaldoCuota) - MoraACuenta
'
'    If CalculoMora < 0 Then CalculoMora = 0
    
End Function

'----------------------------------------------------------------------------------------------------------------------
'   Retorna el nombre del icono segun el estado de la factura
'----------------------------------------------------------------------------------------------------------------------
Public Function IconoDeVencimiento(Vencimiento As Date, Optional TipoDelCredito As Integer = TipoCredito.Normal)

Dim aIcono As String

    Select Case TipoDelCredito
        Case TipoCredito.Gestor: aIcono = "Gestor"
        Case TipoCredito.Incobrable: aIcono = "Perdida"
        Case TipoCredito.Clearing: aIcono = "Clearing"
    
        Case TipoCredito.Normal
            Select Case Vencimiento - gFechaServidor
                Case Is < 0
                    
                    Select Case Abs(Vencimiento - gFechaServidor)
                        Case Is < paToleranciaMora: aIcono = "Alerta"
                        Case Is < paIconoVencimientoN2Dias: aIcono = "Vencida"
                        Case Else: aIcono = "No"
                    End Select
                    
                Case Is >= 0
                    Select Case Abs(Vencimiento - gFechaServidor)
                        Case Is < paIconoPendienteN2Dias: aIcono = "Si"
                        Case Else: aIcono = "Blanco"
                    End Select
            End Select
    End Select
    
    IconoDeVencimiento = aIcono
    
End Function

'---------------------------------------------------------------------------------------------------
'   En base al cumplimiento calcula el puntaje del credito
'---------------------------------------------------------------------------------------------------
Public Function FormatoPuntaje(Cumplimiento As String) As Integer

Dim aSuma As Integer
Const sCumplimiento = "0123456789ABCDEFGHIJKLMNOPQRSTX"
Dim iI As Integer

    Cumplimiento = Trim(Cumplimiento)
    For iI = 1 To Len(Cumplimiento)
        aSuma = aSuma + InStr(sCumplimiento, Mid(Cumplimiento, iI, 1)) - 1
    Next iI
    
    FormatoPuntaje = CInt(aSuma / Len(Cumplimiento))
    
End Function

