Attribute VB_Name = "modFechas"
'-------------------------------------------------------------------------------------------------
' ModFechas.
'
' Este modulo contiene rutinas que se utilizan para resultados distintos de fechas.
'
'Autores:
'   A&A analistas ......   graduados un 4/5/98
'   Junio 1998
'-------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------
'   Funciones:
'       UltimoDia(Fecha As Date)
'       PrimerDia(Fecha As Date)
'       SumoDias(Fecha As String, Dias As Long)
'       RestoDias(Fecha As String, Dias As Long)
'       RestoFechas(f1 As String, f2 As String)
'       RestoPeriodo(f1 As String, Per As Integer)     --- Per en meses
'       ValidoPeriodoFechas(Cadena As String)
'------------------------------------------------------------------------------------------------

Option Explicit
Public gFechaServidor As Date
'----------------------------------------------------------------------------------------------------
'   Consulta por la fecha del servidor y la carga en la variable global gFechaServidor
'----------------------------------------------------------------------------------------------------
Public Sub FechaDelServidor()

    Dim RsF As rdoResultset
    
    On Error GoTo errFecha
    Cons = "Select GetDate()"
    Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    gFechaServidor = RsF(0)
    RsF.Close
    
    On Error Resume Next
    Date = gFechaServidor
    Time = gFechaServidor
    Exit Sub

errFecha:
    gFechaServidor = Now
End Sub

'-----------------------------------------------------------------------------------
' Retorna el primer día del mes de la fecha pasada como parametro.
'-----------------------------------------------------------------------------------
Public Function PrimerDia(fecha As Date)
    PrimerDia = fecha - Format(fecha, "dd") + 1
End Function

'------------------------------------------------------------------------------------
' Retorna el último día del mes de la fecha pasada como parametro.
'------------------------------------------------------------------------------------
Public Function UltimoDia(fecha As Date)
    UltimoDia = DateAdd("m", 1, fecha)
    UltimoDia = UltimoDia - Format(UltimoDia, "dd")
End Function

'------------------------------------------------------------------------------
' Retorna la cantidad de meses que resulta en restar dos fechas.
' OBSERVACIONES:
'  El PRIMER PARAMETRO debe ser la FECHA MAS CHICA.
' EJ. '01/11/97' - '01/12/97' = 01
'---------------------------------
Public Function RestoFechas(f1 As String, f2 As String)
    RestoFechas = DateDiff("m", f1, f2)
End Function

'-------------------------------------------------------------------------------------
'   Funcion que retorna el día exacto al restar meses a una fecha.
'   Ej: A una Fecha dd/mm/yyy le resto 5 meses ------> dd/mm/yyyy
'------------------------------------------------------------------------------------
Public Function RestoPeriodo(f1 As String, Per As Integer)
    RestoPeriodo = DateAdd("m", Per * -1, f1)
End Function

'-----------------------------------------------------------------------------------
'   Resta a Fecha la cantidad de días Dias
'   Retorna: un string con la fecha resultado formato dd/mm/yy
'-----------------------------------------------------------------------------------
Public Function RestoDias(fecha As String, Dias As Long)

    RestoDias = Format(CDate(fecha) - Dias, "dd/mm/yyyy")
    
End Function

'-----------------------------------------------------------------------------------
'   Resta a Fecha la cantidad de días Dias
'   Retorna: un string con la fecha resultado formato dd/mm/yy
'-----------------------------------------------------------------------------------
Public Function SumoDias(fecha As String, Dias As Long)

    SumoDias = Format(CDate(fecha) + Dias, "dd/mm/yyyy")
    
End Function

'-----------------------------------------------------------------------------------------------------
'   Valida un periodo de fechas pasado como string. Se valida:
'       Fecha                    Iguales a una fecha
'       >Fecha                  Mayores a una fecha
'       <Fecha                  Menores a una fecha
'       EFechaYFecha        Entre fecha Y Fecha
'
'   Retorna: String NULO si no son fechas, o sea ""
'                String con > o < y la fecha formato dd/mm/yyyy
'                String con una o dos fechas formato dd/mm/yyyy o dd/mm/yyyydd/mm/yyyy
'                       si son dos hay que leer hasta la posicion 10 y de la 11 en adelante
'------------------------------------------------------------------------------------------------------
Public Function ValidoPeriodoFechas(Cadena As String, Optional ConEY As Boolean = False)
    
Dim aS1 As String
Dim aS2 As String

    ValidoPeriodoFechas = ""
    Cadena = UCase(Cadena)
    If IsDate(Cadena) Then
        ValidoPeriodoFechas = Format(Cadena, "dd/mm/yyyy")
        Exit Function
    End If
    
    If Mid(Cadena, 1, 1) = ">" Or Mid(Cadena, 1, 1) = "<" Then
        If IsDate(Mid(Cadena, 2, Len(Cadena))) Then
            ValidoPeriodoFechas = Mid(Cadena, 1, 1) & Format(Mid(Cadena, 2, Len(Cadena)), "dd/mm/yyyy")
            Exit Function
        End If
    End If
    
    If Mid(Cadena, 1, 1) = "E" Then
        If InStr(Cadena, "Y") <> 0 Then
            If IsDate(Mid(Cadena, 2, InStr(Cadena, "Y") - 2)) Then
                aS1 = Mid(Cadena, 2, InStr(Cadena, "Y") - 2)
                If IsDate(Mid(Cadena, InStr(Cadena, "Y") + 1, Len(Cadena))) Then
                    aS2 = Mid(Cadena, InStr(Cadena, "Y") + 1, Len(Cadena))
                    If Not ConEY Then
                        ValidoPeriodoFechas = Format(aS1, "dd/mm/yyyy") & Format(aS2, "dd/mm/yyyy")
                    Else
                        ValidoPeriodoFechas = "E" & Format(aS1, "dd/mm/yyyy") & "Y" & Format(aS2, "dd/mm/yyyy")
                    End If
                    Exit Function
                End If
            End If
        End If
    End If
    
End Function
'-------------------------------------------------
'Le suma a una fecha cierta cant. de meses.
'Ej. '01/01/97' + 01 = 01/02/97
'-------------------------------------------------
Public Function SumoPeriodo(f1 As String, Per As Integer)
    SumoPeriodo = DateAdd("m", Per, f1)
End Function

'---------------------------------------------------------------------------------------------------------------------------
'   Arma la porción de la consulta de las fechas, para procesar el ingreso de un rango de fechas.
'   PARAMETROS:
'       Condicion: Where o And
'       Campo: Nombre del campo fecha en la BD.
'       CadenaFecha: string donde se ingreso la fecha.
'---------------------------------------------------------------------------------------------------------------------------
Public Function ConsultaDeFecha(Condicion As String, Campo As String, CadenaFecha As String) As String

Dim aStr As String

    Condicion = Trim(Condicion)
    Campo = Trim(Campo)
    
    If IsDate(CadenaFecha) Then         'Igual a una Fecha
       aStr = " " & Condicion & " " & Campo _
              & " Between '" & Format(CadenaFecha, "mm/dd/yyyy") & "'" _
              & " And '" & Format(CadenaFecha, "mm/dd/yyyy 23:59") & "'"
    Else
        'Mayor o menor a una fecha
       If Mid(CadenaFecha, 1, 1) = ">" Or Mid(CadenaFecha, 1, 1) = "<" Then
            If Mid(CadenaFecha, 1, 1) = ">" Then
               aStr = " " & Condicion & " " & Campo & " > '" & Format(Mid(CadenaFecha, 2, 10), "mm/dd/yyyy") & "'"
            Else
               aStr = " " & Condicion & " " & Campo & " < '" & Format(Mid(CadenaFecha, 2, 10), "mm/dd/yyyy") & "'"
            End If
       Else
            'Entre una fecha y tal otra
            CadenaFecha = ValidoPeriodoFechas(CadenaFecha)
            aStr = " " & Condicion & " " & Campo _
                   & " Between '" & Format(Mid(CadenaFecha, 1, 10), "mm/dd/yyyy 00:00") & "'" _
                   & " And '" & Format(Mid(CadenaFecha, 11, 10), "mm/dd/yyyy 23:59") & "'"
       End If
    End If
    ConsultaDeFecha = aStr
    
End Function

Public Function RetornoFormatoFechaConsulta(strFecha As String) As String
    
    If InStr(strFecha, ">") > 0 Or InStr(strFecha, "<") > 0 Then
        If InStr(strFecha, "=") = 0 Then
            If IsDate(Mid(Trim(strFecha), 2, Len(strFecha))) Then RetornoFormatoFechaConsulta = Mid(strFecha, 1, 1) & "'" & Format(Mid(strFecha, 2, Len(strFecha)), "mm/dd/yyyy") & "'"
        ElseIf InStr(strFecha, "=") = 2 Then
            If IsDate(Mid(Trim(strFecha), 3, Len(strFecha))) Then
                RetornoFormatoFechaConsulta = Mid(strFecha, 1, 2) & "'" & Format(Mid(strFecha, 3, Len(strFecha)), "mm/dd/yyyy") & "'"
            End If
        End If
    Else
        If InStr(strFecha, "=") = 0 Then
            If IsDate(Trim(strFecha)) Then RetornoFormatoFechaConsulta = "=" & "'" & Format(strFecha, "mm/dd/yyyy") & "'"
        Else
            If IsDate(Mid(Trim(strFecha), 2, Len(strFecha))) Then RetornoFormatoFechaConsulta = Mid(strFecha, 1, 1) & "'" & Format(Mid(strFecha, 2, Len(strFecha)), "mm/dd/yyyy") & "'"
        End If
    End If

End Function
