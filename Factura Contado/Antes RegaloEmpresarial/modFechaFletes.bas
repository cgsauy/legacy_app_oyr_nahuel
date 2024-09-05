Attribute VB_Name = "modFechaTipoFlete"
Option Explicit

Public Sub ValidoFechaFleteVtaWeb(ByVal sIDsEnvio As String)
On Error GoTo errVal
Dim sQy As String, Agenda As String, AgendaAbierta As String
Dim fechaAge As Date, fEnvio As Date

    If sIDsEnvio = "" Then Exit Sub

    sQy = "SELECT IsNull(TFlAgenda, 0) as Agenda, IsNull(TFlAgendaHabilitada, 0) as AgendaH, IsNull(TFLFechaAgeHab, GetDate()) as FAgenda, EnvFechaPrometida" & _
        " FROM TipoFlete INNER JOIN Envio ON EnvTipoFlete = TFlCodigo AND EnvCodigo IN (" & sIDsEnvio & ")" & _
        " LEFT OUTER JOIN TipoHorario ON TFlRangoHs = THoID" & _
        " ORDER BY TFlDescripcion"

Dim rsT As rdoResultset
    Set rsT = cBase.OpenResultset(sQy, rdOpenForwardOnly, rdConcurValues)
    If Not rsT.EOF Then
        fechaAge = rsT("FAgenda")
        Agenda = rsT("Agenda")
        AgendaAbierta = rsT("AgendaH")
        If Not IsNull(rsT("EnvFechaPrometida")) Then fEnvio = rsT("EnvFechaPrometida") Else fEnvio = Date
    Else
        rsT.Close
        MsgBox "No se logro controlar la fecha de envío, verifique.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    rsT.Close

Dim dAux As Date
    If fechaAge < Date Then dAux = Date Else dAux = fechaAge
Dim Matriz As String
    If DateDiff("d", fechaAge, Date) >= 7 Then
        'Como cerro hace una semana tomo la agenda normal.
        Matriz = superp_MatrizSuperposicion(Agenda)
    Else
        Matriz = superp_MatrizSuperposicion(AgendaAbierta)
    End If

Dim iSuma As Integer
    If Matriz <> "" Then
        If BuscoProximoDia(fEnvio, Matriz) <> 0 Then
            'NO es abierto.
            'Busco el primer día disponible.
            iSuma = BuscoProximoDia(dAux, Matriz)
            If iSuma <> -1 Then
                fEnvio = Format(DateAdd("d", iSuma, dAux), "dd/mm/yyyy")
                sQy = "UPDATE Envio SET EnvFechaPrometida = '" & Format(fEnvio, "yyyyMMdd") & "' WHERE EnvCodigo IN (" & sIDsEnvio & ")"
                cBase.Execute sQy
                MsgBox "La fecha del envío fue modificada para el primer día disponible.", vbInformation, "CAMBIO EN ENVÍO"
                Exit Sub
            End If
        End If
    End If

    Exit Sub
errVal:
    clsGeneral.OcurrioError "Error al validar la fecha del envío.", Err.Description, "ATENCIÓN"
End Sub

Private Function BuscoProximoDia(dFecha As Date, strMat As String)
Dim rsHora As rdoResultset
Dim intDia As Integer, intSuma As Integer
    
    'Por las dudas que no cumpla en la semana paso la agenda normal.
    
    On Error GoTo errBDER
    BuscoProximoDia = -1
    
    'Consulto en base a la matriz devuelta.
    Cons = "Select * From HorarioFlete Where HFlIndice IN (" & strMat & ")"
    Set rsHora = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsHora.EOF Then
        'Busco el valor que coincida con el dia de hoy y ahí busco para arriba.
        intSuma = 0
        Do While intSuma < 7
            rsHora.MoveFirst
            intDia = Weekday(dFecha + intSuma)
            Do While Not rsHora.EOF
                If rsHora!HFlDiaSemana = intDia Then
                    BuscoProximoDia = intSuma
                    GoTo Encontre
                End If
                rsHora.MoveNext
            Loop
            intSuma = intSuma + 1
        Loop
        rsHora.Close
    End If

Encontre:
    rsHora.Close
    Exit Function
    
errBDER:
    clsGeneral.OcurrioError "Error al buscar el primer día disponible para el tipo de flete.", Trim(Err.Description)
End Function

