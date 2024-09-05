Attribute VB_Name = "modAutonumerico"

Public Enum TAutonumerico
    Envio = 1
    Cliente = 2
    Direccion = 3
    AuxiliarEnvio = 4
    Llamadas = 5
    Estacionamiento = 6
    ClearingConsulta = 7
    EMailDireccion = 8
    Cotizacion = 9
End Enum

Dim idContador As Long

Public Function Autonumerico(Tabla As Integer) As Long
   
    Dim Ok As Boolean, Intentos As Integer
    Intentos = 0: Ok = False
    
    Do While Not Ok And Intentos < 10
        Intentos = Intentos + 1
        Ok = PidoAutonumerico(Tabla)
    Loop
    
    Autonumerico = idContador
    
End Function

Private Function PidoAutonumerico(idTabla As Integer) As Boolean

    On Error GoTo errConcurr
    PidoAutonumerico = False
    Dim RsDoc As rdoResultset
    
    Cons = "Select * from Autonumerico Where Tabla = " & idTabla
    Set RsDoc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsDoc.EOF Then
        RsDoc.AddNew
        RsDoc!Tabla = idTabla
        RsDoc!Contador = 1
        idContador = 1
        RsDoc.Update
    Else
        idContador = RsDoc!Contador + 1
        If idTabla = TAutonumerico.AuxiliarEnvio And idContador > 30000 Then idContador = 1
        RsDoc.Edit
        RsDoc!Contador = idContador
        RsDoc.Update
    End If
    RsDoc.Close
    
    PidoAutonumerico = True
    Exit Function
    
errConcurr:
    RsDoc.Close
End Function
